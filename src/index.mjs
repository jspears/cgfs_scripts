import date from 'date-and-time';
import * as XLSX from 'xlsx/xlsx.mjs';
/* load 'fs' for readFile and writeFile support */
import * as fs from 'fs';
import esMain from 'es-main';

/* load the codepage support library for extended support with older formats  */
//import * as cpexcel from 'xlsx/dist/cpexcel.full.mjs';
//XLSX.set_cptable(cpexcel);
XLSX.set_fs(fs);

const COLUMNS = 'Start_Date	Start_Time	End_Date	End_Time	Title	Description	Location	Location_URL	Location_Details	All_Day_Event	Event_Type	Tags	Team1_ID	Team1_Division_ID	Team1_Is_Home	Team2_ID	Team2_Division_ID	Team2_Name	Custom_Opponent	Event_ID	Game_ID	Affects_Standings	Points_Win	Points_Loss	Points_Tie	Points_OT_Win	Points_OT_Loss	Division_Override'.split('\t');

const cleanLoc = v => v?.toLowerCase().replace(/\([^)]*\)/, '').replaceAll(/Field|Park|Elementary/ig, '').replaceAll(/[-_,.#]+?|\s+?/g, ' ').replaceAll(/\s{2,}/g, ' ').trim();


const isCgfs = v => /cgfs/i.test(v);

const fixTime = (time) => {
  const [_, h, m, a ] = /(\d{1,2})\:(\d{1,2})\s*(am|pm)?/i.exec(time) || [];
  if (!_) {
    throw new Error(`unknown time "` + time + '" ' + _);
  }

  return ( a && h < 12 && a.toUpperCase() == 'pm') ? (12 + h) : h + ':' + m;
};

const isValidDate = d => d && !isNaN(d.getTime());

const parseDate = (obj, after=8, before = 20)=>{
  const d = _parseDate(obj);
  if (!d){
    return d;
  }
  if (after && d.getHours() < after){
    throw new Error(`Invalid time must be after ${after} ${d} ${d.getHours()}`);
  }
  if (before && d.getHours() > before){
    throw new Error( `Invalid time must be before ${before} ${d} ${d.getHours()}`);
  }

  return d;
}

const _parseDate = (obj) => {

  const str = (obj.Date.split(' ')[0] + ' ' + fixTime(obj.Time)).trim() ;
  const newDate = date.parse(str, 'M/D/YYYY H:mm');
  if (!isValidDate(newDate)) {
    throw new Error(`Invalid Date "${newDate}" "${str}"` + obj);
  }
  
  return newDate;

};


const parseAge = (str) => (/(\d{1,2})U/.exec(str) || [])[1];

const parseFile = (file) => {
  const workbook = XLSX.readFile(file, {});
  const schedules = workbook.Sheets['SCHEDULE'] ? [[workbook.Sheets['SCHEDULE'], parseAge(file)]] : workbook.SheetNames.filter(v => /schedule/i.test(v)).map(v =>
    [workbook.Sheets[v], parseAge(v)]
  );
  const fieldSheet = workbook.Sheets['Field Information'];

  const fieldsArr = XLSX.utils.sheet_to_json(fieldSheet);
  fieldsArr.reduce((ret, v) => v['League'] ?? (v['League'] = ret), '');
  fieldsArr.reduce((ret, v) => v['Field Address'] ?? (v['Field Address'] = ret), '');

  const fields = fieldsArr.reduce((ret, v) => {
    const fName = v['Field Name'] ?? v['Field'];

    ret[cleanLoc(fName)] = v;
    v['Field Address'] = v['Field Address']?.replace(/\r\n/g, ',')?.trim();
    if (v['Other Info']) {
      v['Field Name'] = `${fName} (${v['Other Info']})`;
    }
    return ret;
  }, {});

  return schedules.reduce((ret, [sheet, age]) => {
    const resp = parseSchedule(sheet, age, fields);
    ret.push(...resp);
    return ret;
  }, []);
};

const parseSchedule = (sheet, age, fields) => {
  const aschedule = XLSX.utils.sheet_to_json(sheet, { dateNF: false, raw: false });

  aschedule.reduce((ret, obj) => {
    obj._age = age;
    if (obj.Time) {
      if (!obj.Date) {
        obj.Date = ret;
      }
        obj.DateTime = parseDate(obj);
    }
    return obj.Date || ret;
  }, null);


  const schedule = aschedule.filter(v => (v && v['Home Team'] && v['Away Team']));

  const resolveLocation = (location) => {
    const field = fields[cleanLoc(location)];
    if (!field) {
      throw Error(`could not find field for ${JSON.stringify(location)} '${cleanLoc(location)}'`);
    }
    return field;
  };
  const teamId = v => age + v.replaceAll(/\s+?/g, '').trim().toUpperCase();

  return schedule.map(v => {

    const home = teamId(v['Home Team']);
    const away = teamId(v['Away Team']);
    const isAway = !isCgfs(home);

    if (!isCgfs(away) && isAway) {
      return null;
    }
    
    const end = new Date(v.DateTime.getTime() + 2 * 3600 * 1000);
    const loc = resolveLocation(v.Location || v['Field Name'] || v['Field']);

    const val = ({
      Start_Date: date.format(v.DateTime, 'M/D/YY'),
      Start_Time: date.format(v.DateTime, 'H:mm'),
      End_Date: date.format(end, 'M/D/YY'),
      End_Time: date.format(end, 'H:mm'),
      Location: loc['Field Name'],
      Location_Details: loc['Field Address'],
      Event_Type: 'Game',
      Team1_ID: isAway ? away : home,
      Team1_Is_Home: isAway ? 0 : 1,
      ...(isCgfs(home) && isCgfs(away) ? {
        Team2_ID: isAway ? home : away
      } : {
        Team2_ID: isAway ? home : away,
        Team2_Name: v[isAway ? 'Home Team' : 'Away Team'],
        Custom_Opponent: `TRUE`,
      })
    });
    return val;

  }).filter(Boolean);
};

const quote = v => {
  if (v == null) {
    return '';
  };
  if (/^[\w-_+:/]+?$/.test(v)) {
    return v;
  }
  return JSON.stringify(v);
};

const toCSV = (objs) => objs.reduce((ret, o) => (ret + COLUMNS.map(v => quote(o[v])).join(',') + '\n'), '');

export function main(files) {
  console.warn(`processing ${files}`);
  console.log(files.reduce((ret, name) => `${ret}${toCSV(parseFile(name))}`, COLUMNS.join(',') + '\n'));
 
}

if (esMain(import.meta)) {
  main(process.argv.slice(2));
}
