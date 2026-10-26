// A provider week in the layout of a real Castro IVK template (Tuesday to Friday), under an invented provider. Test data only.
const W = {
  2: ["8:30 AM Admin", "9:00 AM New Patient", "9:30 AM New Patient", "10:00 AM Interventional Consult", "10:30 AM Interventional Consult", "11:00 AM Telepsych", "11:30 AM Meeting",
    "12:00 PM Lunch", "12:30 PM IV Ketamine Consult", "1:00 PM IV Ketamine Consult", "1:30 PM TMS Follow Up", "2:00 PM New Patient", "2:30 PM New Patient",
    "3:00 PM Interventional Consult", "3:30 PM Interventional Consult", "4:00 PM Open Schedule", "4:30 PM Telepsych"],
  3: ["8:30 AM Admin", "9:00 AM IV Ketamine Start/Index", "9:30 AM IV Ketamine Start/Index", "10:00 AM IV Ketamine Follow Up", "10:30 AM Flex Post Check", "11:00 AM Flex Post Check",
    "11:30 AM Admin", "12:00 PM Lunch", "12:30 PM IV Ketamine Start/Index", "1:00 PM IV Ketamine Follow Up", "1:30 PM IV Ketamine Follow Up", "2:00 PM IVK Post Check",
    "2:30 PM IV Ketamine Start/Index", "3:00 PM IV Ketamine Start/Index", "3:30 PM TMS Follow Up", "4:00 PM IVK Post Check", "4:30 PM Flex Post Check", "5:00 PM Telepsych"],
  4: ["8:30 AM Admin", "9:00 AM IV Ketamine Start/Index", "9:30 AM IV Ketamine Follow Up", "10:00 AM IV Ketamine Follow Up", "10:30 AM Flex Post Check", "11:00 AM Meeting",
    "11:30 AM Meeting", "12:00 PM IV Ketamine Follow Up", "12:30 PM Telepsych", "1:00 PM Telepsych", "1:30 PM Flex Post Check", "2:00 PM New Patient", "2:30 PM New Patient",
    "3:00 PM TMS Follow Up", "3:30 PM Interventional Consult", "4:00 PM Interventional Consult", "4:30 PM Telepsych"],
  5: ["8:30 AM Admin", "9:00 AM IV Ketamine Start/Index", "9:30 AM IV Ketamine Start/Index", "10:00 AM IV Ketamine Follow Up", "10:30 AM Flex Post Check", "11:00 AM Flex Post Check",
    "11:30 AM Admin", "12:00 PM Lunch", "12:30 PM IV Ketamine Start/Index", "1:00 PM IV Ketamine Follow Up", "1:30 PM IV Ketamine Follow Up", "2:00 PM IVK Post Check",
    "2:30 PM IV Ketamine Start/Index", "3:00 PM IV Ketamine Start/Index", "3:30 PM TMS Follow Up", "4:00 PM IVK Post Check", "4:30 PM Flex Post Check"]
};
/* day -> Map("day|minutes" -> type) for template.layout's copy option */
function copyMap(TM) {
  const m = new Map();
  for (const [d, L] of Object.entries(W)) for (const s of L) { const i = s.indexOf("M ") + 2; m.set(`${d}|${TM.toMin(s.slice(0, i - 1))}`, s.slice(i)); }
  return m;
}
/* the layout written into a sheet-shaped array of values (formulas as text), as excel.js would put it on the IVK Template tab */
function sheet(TM, grids) {
  const rows = [];
  for (const lay of grids) for (const [r, c, v] of lay.put) { while (rows.length <= r) rows.push(Array(16).fill("")); rows[r][c] = v; }
  return rows;
}
module.exports = { W, copyMap, sheet };
