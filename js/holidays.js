// General national holiday dates only; substitute days and individual entitlements are not included.
// Sources and excluded substitute dates: rules/HOLIDAYS_2026_2027.md.
export const CALENDAR_VERSION = 'TW-NATIONAL-2026Q4-2027-v1（2026-10-19 核對）';
export const CALENDAR_SOURCES = {
  law:'https://law.moj.gov.tw/LawClass/LawAll.aspx?pcode=D0020095',
  2026:'https://www.dgpa.gov.tw/information?pid=12574&uid=82',
  2027:'https://www.dgpa.gov.tw/information?pid=13114&uid=30'
};
export const NATIONAL_HOLIDAYS = {
  '2026-10':[['10','國慶日'],['25','臺灣光復暨金門古寧頭大捷紀念日']],
  '2026-11':[],
  '2026-12':[['25','行憲紀念日']],
  '2027-01':[['01','中華民國開國紀念日']],
  '2027-02':[['04','小年夜'],['05','除夕'],['06','春節初一'],['07','春節初二'],['08','春節初三'],['28','和平紀念日']],
  '2027-03':[],
  '2027-04':[['04','兒童節'],['05','清明節']],
  '2027-05':[['01','勞動節']],
  '2027-06':[['09','端午節']],
  '2027-07':[],
  '2027-08':[],
  '2027-09':[['15','中秋節'],['28','孔子誕辰紀念日／教師節']],
  '2027-10':[['10','國慶日'],['25','臺灣光復暨金門古寧頭大捷紀念日']],
  '2027-11':[],
  '2027-12':[['25','行憲紀念日']]
};
export const DEFAULT_HOLIDAYS = Object.fromEntries(Object.entries(NATIONAL_HOLIDAYS)
  .map(([month,days])=>[month,days.map(([day])=>`${month}-${day}`)]));
