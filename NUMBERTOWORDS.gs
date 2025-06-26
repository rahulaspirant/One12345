/**
* Converts a number to its English word representation.
*
* 
@param
 {number} value The number to convert.
* 
@return
 {string} The number in words, or "Invalid Input" if the input is not a number.
* 
@customfunction

*/
function NUMBERTOWORDS(value) {
if (typeof value === "number") {
return numberToWords(value);
}
return "Invalid Input";
}

function numberToWords(num) {
if (num === 0) return 'Zero';

var belowTwenty = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
var tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
var thousands = ["", "Thousand", "Million", "Billion"];

function helper(n) {
if (n === 0) return "";
else if (n < 20) return belowTwenty[n] + " ";
else if (n < 100) return tens[Math.floor(n / 10)] + " " + helper(n % 10);
else return belowTwenty[Math.floor(n / 100)] + " Hundred " + helper(n % 100);
}

let word = "";
let i = 0;

while (num > 0) {
if (num % 1000 !== 0) {
word = helper(num % 1000) + thousands[i] + " " + word;
}
num = Math.floor(num / 1000);
i++;
}

return word.trim();
}
