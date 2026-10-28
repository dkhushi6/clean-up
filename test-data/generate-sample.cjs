const XLSX = require("xlsx");

// Header uses variant names on purpose, to test fuzzy column matching,
// plus one unrelated extra column to test "ignore + mention" handling.
const headers = ["Cust Name", "Mobile No", "Amt", "Branch", "Loan Date", "Status", "Remarks"];

const rows = [
  // --- Clean rows ---
  ["Ravi Kumar", "9876501234", "500000", "Pune", "15/03/2024", "Active", "ok"],
  ["Priya Singh", "9823456780", "2,50,000", "Mumbai", "21-06-2023", "Closed", "ok"],
  ["Mohammed Imran", "7012345678", "7L", "Delhi", "2023-11-20", "NPA", "ok"],

  // --- Fully blank row (should be skipped) ---
  ["", "", "", "", "", "", ""],

  // --- branch checks ---
  ["Sunita Rao", "9988001122", "300000", "", "20/02/2024", "Active", "blank branch"],

  // --- customer_name checks ---
  ["", "9988002233", "400000", "Pune", "24/02/2024", "Active", "blank name"],
  ["NA", "9988003344", "150000", "Chennai", "25/01/2024", "Closed", "junk name"],
  ["Amit123", "9988004455", "600000", "Mumbai", "18/03/2024", "Active", "name has digits"],
  ["X", "9988005566", "250000", "Delhi", "22/01/2024", "Active", "single-char name"],

  // --- loan_amount checks ---
  ["Deepak Verma", "9988006677", "", "Pune", "14/02/2024", "Active", "blank amount"],
  ["Kavita Joshi", "9988007788", "abc", "Mumbai", "19/02/2024", "Active", "non-numeric amount"],
  ["Sanjay Gupta", "9988008899", "-5000", "Delhi", "21/02/2024", "Active", "negative amount"],
  ["Meena Iyer", "9988009900", "1,50,00,000", "Chennai", "25/02/2024", "Active", "exceeds 1 crore"],

  // --- loan_date checks ---
  ["Rajesh Nair", "9988010011", "350000", "Pune", "", "Active", "blank date"],
  ["Shalini Reddy", "9988011122", "450000", "Mumbai", "32/13/2024", "Active", "bad date format"],
  ["Vikram Shah", "9988012233", "275000", "Delhi", "03/04/2024", "Active", "ambiguous date"],
  ["Pooja Mehta", "9988013344", "325000", "Chennai", "15/12/2026", "Active", "future date"],
  ["Arun Das", "9988014455", "180000", "Pune", "20/05/1998", "Closed", "too old date"],

  // --- phone checks ---
  ["Neha Kapoor", "", "220000", "Mumbai", "28/03/2024", "Active", "blank phone"],
  ["Suresh Pillai", "998801556", "310000", "Delhi", "21/03/2024", "Active", "9-digit phone"],
  ["Anjali Bose", "5988016677", "290000", "Chennai", "13/03/2024", "Active", "starts with 5"],
  ["Farhan Khan", "9999999999", "260000", "Pune", "16/03/2024", "Active", "fake repeated phone"],
  ["Geeta Menon", "1234567890", "330000", "Mumbai", "17/03/2024", "Active", "fake sequential phone"],

  // --- status checks ---
  ["Rohit Malhotra", "9988019900", "410000", "Delhi", "20/03/2024", "", "blank status"],
  ["Divya Krishnan", "9988020011", "190000", "Chennai", "23/03/2024", "Pending", "unmappable status"],

  // --- row-level checks ---
  ["Alok Tiwari", "9988021122", "500000", "Pune", "20/10/2026", "Closed", "closed but recent date"],
  ["Sneha Pandey", "", "", "", "21/01/2024", "", "incomplete row, 3+ blanks"],
  ["Imran Qureshi", "9876543210", "300000", "Mumbai", "22/04/2024", "Active", "amount/phone swapped"],

  // --- cross-row checks ---
  ["Sameer Joshi", "9988023344", "500000", "Pune", "21/02/2024", "Active", "original for exact dup"],
  ["Sameer Joshi", "9988023344", "500000", "Pune", "21/02/2024", "Active", "exact duplicate of row above"],
  ["Kiran Thakur", "9988024455", "275000", "Delhi", "25/02/2024", "Active", "same phone as next, diff name"],
  ["Nisha Bhatt", "9988024455", "310000", "Chennai", "26/02/2024", "Active", "same phone as row above"],
  ["Wasim Ali", "9988025566", "450000", "Pune", "20/02/2024", "Active", "same name+amount as next"],
  ["Wasim Ali", "9988026677", "450000", "Mumbai", "23/02/2024", "Closed", "same name+amount as row above"],
  ["Lata Nambiar", "9988027788", "380000", "Delhi", "14/02/2024", "Active", "same name+date as next"],
  ["Lata Nambiar", "9988028899", "220000", "Chennai", "14/02/2024", "Active", "same name+date as row above"],

  // --- suggestion checks ---
  // Phone blank, but an otherwise-identical row has one: suggest it (also a cross-row duplicate flag).
  ["Ritu Saxena", "9988029900", "325000", "Pune", "17/02/2024", "Active", "phone source for suggestion"],
  ["Ritu Saxena", "", "325000", "Pune", "17/02/2024", "Active", "blank phone - should suggest from row above"],
];

// Swap amount/phone for the "swapped fields" test row (Imran Qureshi):
// put a phone-looking value in the amount column and an amount-looking value in the phone column.
const swapIdx = rows.findIndex((r) => r[6] === "amount/phone swapped");
rows[swapIdx][2] = "9876501234"; // amount column holds a phone-shaped number
rows[swapIdx][1] = "₹3,00,000"; // phone column holds a currency-shaped value

// Make "Pune" the overwhelming majority branch (keeping intentional blanks),
// so the blank-branch row (Sunita Rao) gets a one-click "Pune?" suggestion.
for (const r of rows) {
  if (r[3] !== "") r[3] = "Pune";
}

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
XLSX.utils.book_append_sheet(wb, ws, "Loans");
XLSX.writeFile(wb, __dirname + "/sample-loans.xlsx");

console.log(`Wrote ${rows.length} data rows to sample-loans.xlsx`);
