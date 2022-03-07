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


const isCgfs = v => /cgfs/i.test(v);

const fixTime = (time) => {
  const [_, t, a = 'PM'] = /(\d{1,2}\:\d{1,2})\s*(am|pm)?/i.exec(time) || [];
  if (!_) {
    throw new Error(`unknown time "` + time + '" ' + _);
  }
  return t + ' ' + (a.toUpperCase());
};


const parseDate = (obj) => {

  const str = (obj.Date.split(' ')[0] + ' ' + fixTime(obj.Time)).trim();
  const newDate = date.parse(str, 'M/D/YYYY h:m A');
  if (isNaN(newDate.getTime())) {
    throw new Error(`Invalid Date "${newDate}" "${str}"` + obj);
  }
  return newDate;

};
const parseAge = (str)=>(/(\d{1,2})U/.exec(str) || [])[1];

const parseFile = (file) => {
  const workbook = XLSX.readFile(file, {});
  const schedules = workbook.Sheets['SCHEDULE'] ? [[workbook.Sheets['SCHEDULE'], parseAge(file)]] : workbook.SheetNames.filter(v=>/schedule/i.test(v)).map(v=>
      [workbook.Sheets[v], parseAge(v)]
 );
 const fieldSheet = workbook.Sheets['Field Information'];

 const fieldsArr = XLSX.utils.sheet_to_json(fieldSheet);
 fieldsArr?.reduce((ret, v) => v['League'] || (v['League'] = ret));

 const fields = fieldsArr.reduce((ret, v) => {
   ret[v['Field Name']?.trim()] = v;
   v['Field Address'] = v['Field Address']?.replace(/\r\n/g, ',')?.trim();
   return ret;
 }, {});


  schedules.forEach(([sheet, age])=>parseSchedule(sheet,age, fields));
};

const parseSchedule = (sheet, age, fields)=>{
//  console.log('age', age, sheet);

//  console.log(`schedule`,  XLSX.utils.sheet_to_json(sheet, { dateNF: false, raw: true }));

  const aschedule = XLSX.utils.sheet_to_json(sheet, { dateNF: false, raw: false })

  aschedule.reduce((ret, obj) => {
    if (obj.Time) {
      obj.Date = ret;
      try {
        obj.DateTime = parseDate(obj);
      } catch (e) {
        console.log(`could not parse `, e, obj);
      }
    } 
    return obj.Date || ret;
  }, null);

  const schedule = aschedule.filter(v => (v.Time && v['Home Team'] && v[
    'Away Team'
  ]));

  const findLocation = (location) => fields[location?.trim()]?.['Field Address'];
  const resolveLocation = (location) => findLocation(location) ?? findLocation(location.split(/[,-]/)[0]) ?? findLocation(location.split(/\s+?/)[0]);


  return schedule.map(v => {
    
    const end = new Date(v.DateTime.getTime() + 2 * 3600 * 1000);
    const home = v['Home Team'];
    const away = v['Away Team'];
    const isAway = !isCgfs(home);

    if (!isCgfs(away) && isAway) {
      return null;
    }

    return ({
      Start_Date: date.format(v.DateTime, 'M/D/YY'),
      Start_Time: date.format(v.DateTime, 'H:mm'),
      End_Date: date.format(end, 'M/D/YY'),
      End_Time: date.format(end, 'H:mm'),
      Location: v.Location || v.Field,
      Location_Details: resolveLocation(v.Location || v.Field),
      Event_Type: 'Game',
      Team1_ID: `${age}${isAway ? away : home}`,
      Team1_Is_Home: isAway ? 0 : 1,
      ...(isCgfs(home) && isCgfs(away) ? {
        Team2_ID: `${age}${away}`
      } : {
        Team2_Name: away,
        Custom_Opponent: `TRUE`,
      })
    });
  }).filter(Boolean);
}
const quote = v=>{
  if (v == null){
    return ''
  };
  if (/^[\w-_+:/]+?$/.test(v)){
    return v;
  }
  return JSON.stringify(v);
}

const toCSV = (objs)=>objs.reduce((ret, o)=>{
    return ret+COLUMNS.map(v=>quote(o[v])).join(',')+'\n'
  },'');

export function main(files){
  console.warn(`processing ${files}`);
  console.log(files.reduce((ret, name)=>`${ret}${toCSV(parseFile(name))}`, COLUMNS.join(',')+'\n'));
 }

 if (esMain(import.meta)) {
  main(process.argv.slice(2));
 }
//console.log(toCSV(parseFile('./sheets/2022 FINAL 10U Interleague Schedule Spring.xlsx')));
