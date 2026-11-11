// Function catalog for the Insert Function dialog, formula autocomplete and argument hints.
// Entries are stored as a compact text table and parsed once at module load:
//   #Category              starts a category block
//   NAME|about|args        args separated by ";"
//   arg spec: [?]name[+][:about]   "?" = optional, "+" = repeating (implies optional)
// An arg without ":about" takes its description from COMMON_ARGS.

export type FnCategory =
  | 'Financial' | 'Date & Time' | 'Math & Trig' | 'Statistical' | 'Lookup & Reference' | 'Database'
  | 'Text' | 'Logical' | 'Information' | 'Engineering' | 'Cube' | 'Web' | 'Compatibility';

export interface FnArg { name: string; optional?: boolean; repeat?: boolean; about: string }
export interface FnInfo { name: string; category: FnCategory; about: string; args: FnArg[]; returns?: string }

const COMMON_ARGS: Readonly<Record<string, string>> = {
  number: 'The number to work with',
  number1: 'First number, cell or range',
  number2: 'More numbers, cells or ranges',
  value: 'The value to check or convert',
  value1: 'First value, cell or range',
  value2: 'More values, cells or ranges',
  text: 'The text to work with',
  text1: 'First text item or range',
  text2: 'More text items or ranges',
  reference: 'A cell or range reference',
  ref1: 'First range to include',
  ref2: 'More ranges to include',
  array: 'A range or array of values',
  array1: 'First range or array',
  array2: 'Second (or further) range or array',
  logical: 'A value or test that is TRUE or FALSE',
  logical1: 'First condition to check',
  logical2: 'More conditions to check',
  rate: 'Interest rate per period',
  nper: 'Total number of payment periods',
  pmt: 'Payment made every period',
  pv: 'Present value: what the payments are worth now',
  fv: 'Future value wanted after the last payment (default 0)',
  type: 'When payments fall due: 0 = end of period, 1 = start',
  per: 'The period to look at, from 1 to nper',
  guess: 'Starting estimate for the result',
  values: 'Range of cash flows (needs at least one negative and one positive)',
  settlement: 'Date the buyer takes ownership of the security',
  maturity: 'Date the security expires',
  issue: 'Date the security was issued',
  basis: 'Day count convention code from 0 to 4 (default 0, US 30/360)',
  frequency: 'Coupon payments per year: 1, 2 or 4',
  redemption: 'Amount repaid per $100 of face value',
  pr: 'Price per $100 of face value',
  yld: 'Annual yield of the security',
  discount: 'Discount rate of the security',
  investment: 'Amount invested in the security',
  cost: 'Starting cost of the asset',
  salvage: 'Value left at the end of the asset life',
  life: 'Number of periods the asset is used for',
  period: 'Period to calculate depreciation for',
  factor: 'Rate of balance decline (default 2, double-declining)',
  start_period: 'First period of the range',
  end_period: 'Last period of the range',
  start_date: 'Date the span begins',
  end_date: 'Date the span ends',
  holidays: 'Optional list of dates to treat as days off',
  weekend: 'Weekend code (1-17) or a 7-character mask like "0000011"',
  serial_number: 'The date or time value to read from',
  year: 'Year number, e.g. 2026',
  month: 'Month number (values past 12 roll over)',
  day: 'Day of the month (values past the month end roll over)',
  hour: 'Hours (0-23)',
  minute: 'Minutes (0-59)',
  second: 'Seconds (0-59)',
  return_type: 'Numbering scheme to use for the result',
  months: 'Months to move; negative goes back in time',
  days: 'Working days to move; negative goes back',
  x: 'The value at which to evaluate the function',
  n: 'Order of the function',
  y: 'Value to apply the inverse to',
  z: 'Standard score to evaluate',
  mean: 'Average of the distribution',
  standard_dev: 'Standard deviation of the distribution',
  cumulative: 'TRUE for the cumulative result, FALSE for the density/mass',
  probability: 'Probability between 0 and 1',
  alpha: 'First shape parameter of the distribution',
  beta: 'Second shape parameter of the distribution',
  A: 'Lower bound of the interval (default 0)',
  B: 'Upper bound of the interval (default 1)',
  deg_freedom: 'Degrees of freedom',
  deg_freedom1: 'Numerator degrees of freedom',
  deg_freedom2: 'Denominator degrees of freedom',
  number_s: 'Number of successes',
  number_f: 'Number of failures',
  trials: 'Number of independent trials',
  probability_s: 'Chance of success on each trial',
  sample_s: 'Successes in the sample',
  number_sample: 'Size of the sample',
  population_s: 'Successes in the population',
  number_pop: 'Size of the population',
  k: 'Position or fraction to return',
  quart: 'Which quartile: 0 to 4',
  tails: '1 for a one-tailed test, 2 for two-tailed',
  sigma: 'Known population standard deviation (optional)',
  size: 'Sample size',
  "known_y's": 'Observed dependent values',
  "known_x's": 'Observed independent values',
  "new_x's": 'X values to calculate predictions for',
  const: 'FALSE forces the constant term to zero or one',
  stats: 'TRUE to also return regression statistics',
  range: 'Cells to test against the condition',
  criteria: 'Condition a cell must meet, e.g. ">5" or "apples"',
  criteria_range1: 'First range to test',
  criteria1: 'Condition for the first range',
  criteria_range2: 'More ranges to test',
  criteria2: 'Condition for each extra range',
  significance: 'Multiple to round to',
  num_digits: 'Number of decimal places (negative rounds left of the point)',
  num_chars: 'How many characters to take',
  num_bytes: 'How many bytes to take',
  start_num: 'Position of the first character to use (1 = first)',
  find_text: 'The text to look for',
  within_text: 'The text to search in',
  old_text: 'The original text',
  new_text: 'Replacement text',
  instance_num: 'Which occurrence to use (default every one / the first)',
  decimals: 'Digits after the decimal point',
  places: 'Character count to pad the result with leading zeros',
  inumber: 'Complex number such as "3+4i"',
  inumber1: 'First complex number',
  inumber2: 'Second (or further) complex number',
  lookup_value: 'The value to look for',
  lookup_array: 'Range or array to search',
  return_array: 'Range or array to return results from',
  table_array: 'Table to search, with keys in its first column or row',
  range_lookup: 'FALSE for an exact match, TRUE or omitted for approximate',
  match_mode: 'How to match: exact, next smaller, next larger or wildcard',
  search_mode: 'Search direction: first-to-last, last-to-first or binary',
  rows: 'Number of rows',
  columns: 'Number of columns',
  lambda: 'A LAMBDA function to apply',
  pad_with: 'Value used to fill gaps (default #N/A)',
  row_num: 'Row number',
  column_num: 'Column number',
  connection: 'Name of the connection to the cube data model',
  member_expression: 'MDX expression naming a member or tuple',
  set_expression: 'MDX expression or CUBESET cell describing a set',
  caption: 'Text to show in the cell instead of the cube caption',
  case_sensitivity: '0 for case-sensitive (default), 1 to ignore case',
  pattern: 'Regular expression to apply',
  delimiter: 'Text that marks the split point',
  if_not_found: 'Value returned when nothing is found',
  field_headers: 'Whether the source has headers and whether to show them',
  filter_array: 'TRUE/FALSE array choosing which rows to keep',
  sort_order: '1 for ascending, -1 for descending',
  function_num: 'Code picking the calculation, e.g. 9 for SUM',
  angle: 'Angle in radians',
  bottom: 'Smallest whole number to return',
  top: 'Largest whole number to return',
  seasonality: 'Season length; 1 detects it, 0 means none',
  data_completion: '1 fills missing points by averaging, 0 treats them as zero',
  aggregation: 'Code for how to combine values that share a time stamp',
  timeline: 'Evenly spaced dates or numbers matching the values',
  target_date: 'Date or point in time to forecast',
};

const DB_ARGS = 'database:Range holding the list, with column labels in the first row;field:Column to use, by label or position;criteria:Range with labels and the conditions rows must meet';

const DATA = `
#Financial
ACCRINT|Interest accrued on a security that pays interest periodically|issue;first_interest:Date of the first interest payment;settlement;rate:Annual coupon rate;par:Face value of the security;frequency;?basis;?calc_method:TRUE accrues from issue, FALSE from the first interest date
ACCRINTM|Interest accrued on a security that pays all interest at maturity|issue;settlement:Maturity date of the security;rate:Annual coupon rate;par:Face value of the security;?basis
AMORDEGRC|French-style declining depreciation per period, scaled by an asset life coefficient|cost;date_purchased:Date the asset was bought;first_period:End date of the first period;salvage;period;rate:Depreciation rate;?basis
AMORLINC|French-style straight-line depreciation for each accounting period|cost;date_purchased:Date the asset was bought;first_period:End date of the first period;salvage;period;rate:Depreciation rate;?basis
COUPDAYBS|Days from the start of the coupon period up to settlement|settlement;maturity;frequency;?basis
COUPDAYS|Length in days of the coupon period that contains settlement|settlement;maturity;frequency;?basis
COUPDAYSNC|Days from settlement to the next coupon date|settlement;maturity;frequency;?basis
COUPNCD|The first coupon date after settlement|settlement;maturity;frequency;?basis
COUPNUM|How many coupons are paid between settlement and maturity|settlement;maturity;frequency;?basis
COUPPCD|The last coupon date on or before settlement|settlement;maturity;frequency;?basis
CUMIPMT|Total interest paid on a loan across a range of periods|rate;nper;pv;start_period;end_period;type
CUMPRINC|Total principal repaid on a loan across a range of periods|rate;nper;pv;start_period;end_period;type
DB|Depreciation for one period using a fixed declining balance|cost;salvage;life;period;?month:Months in the first year (default 12)
DDB|Depreciation for one period using double-declining balance or a chosen factor|cost;salvage;life;period;?factor
DISC|Discount rate of a security|settlement;maturity;pr;redemption;?basis
DOLLARDE|Turns a price written as a fraction into a decimal number|fractional_dollar:Price written as whole.fraction, e.g. 1.02;fraction:Denominator of the fraction part
DOLLARFR|Turns a decimal price into whole.fraction notation|decimal_dollar:Price as a decimal number;fraction:Denominator to express the fraction in
DURATION|Macaulay duration of a security that pays periodic interest|settlement;maturity;coupon:Annual coupon rate;yld;frequency;?basis
EFFECT|Effective annual rate from a nominal rate and compounding count|nominal_rate:Stated annual interest rate;npery:Compounding periods per year
FV|Future value of an investment with steady payments and a fixed rate|rate;nper;pmt;?pv:Present value (default 0);?type
FVSCHEDULE|Grows a starting amount through a series of different rates|principal:Starting amount;schedule:Rates to apply, in order
INTRATE|Interest rate of a fully invested security|settlement;maturity;investment;redemption:Amount received at maturity;?basis
IPMT|Interest portion of a loan payment in a given period|rate;per;nper;pv;?fv;?type
IRR|Internal rate of return of a series of periodic cash flows|values;?guess
ISPMT|Interest paid in one period of a loan that repays principal evenly|rate;per;nper;pv
MDURATION|Modified duration of a security, assuming $100 par|settlement;maturity;coupon:Annual coupon rate;yld;frequency;?basis
MIRR|Rate of return that factors in borrowing cost and reinvestment rate|values;finance_rate:Rate paid on borrowed money;reinvest_rate:Rate earned when cash is reinvested
NOMINAL|Nominal annual rate from an effective rate|effect_rate:Effective annual rate;npery:Compounding periods per year
NPER|Number of periods a loan or investment needs|rate;pmt;pv;?fv;?type
NPV|Net present value of future periodic cash flows at a discount rate|rate:Discount rate for one period;value1:First cash flow;value2+:More cash flows
ODDFPRICE|Price per $100 of a security whose first period is irregular|settlement;maturity;issue;first_coupon:Date of the first coupon;rate:Annual coupon rate;yld;redemption;frequency;?basis
ODDFYIELD|Yield of a security whose first period is irregular|settlement;maturity;issue;first_coupon:Date of the first coupon;rate:Annual coupon rate;pr;redemption;frequency;?basis
ODDLPRICE|Price per $100 of a security whose last period is irregular|settlement;maturity;last_interest:Date of the last coupon;rate:Annual coupon rate;yld;redemption;frequency;?basis
ODDLYIELD|Yield of a security whose last period is irregular|settlement;maturity;last_interest:Date of the last coupon;rate:Annual coupon rate;pr;redemption;frequency;?basis
PDURATION|Periods an investment needs to grow to a target value|rate;pv:Starting value;fv:Target value
PMT|Payment per period for a loan with fixed payments and rate|rate;nper;pv;?fv;?type
PPMT|Principal portion of a loan payment in a given period|rate;per;nper;pv;?fv;?type
PRICE|Price per $100 of a security that pays periodic interest|settlement;maturity;rate:Annual coupon rate;yld;redemption;frequency;?basis
PRICEDISC|Price per $100 of a discounted security|settlement;maturity;discount;redemption;?basis
PRICEMAT|Price per $100 of a security that pays interest at maturity|settlement;maturity;issue;rate:Interest rate at issue;yld;?basis
PV|Present value of a series of equal future payments|rate;nper;pmt;?fv;?type
RATE|Interest rate per period of an annuity|nper;pmt;pv;?fv;?type;?guess
RECEIVED|Amount paid back at maturity on a fully invested security|settlement;maturity;investment;discount;?basis
RRI|Equivalent steady interest rate for an investment's growth|nper;pv:Starting value;fv:Ending value
SLN|Straight-line depreciation for one period|cost;salvage;life
STOCKHISTORY|Pulls historical price data for a stock or fund|stock:Ticker text or a Stocks data type cell;start_date;?end_date;?interval:0 daily, 1 weekly, 2 monthly;?headers:0 none, 1 show, 2 show with instrument name;?property0:First column to return (0 date, 1 close, ...);?property1+:More columns to return
SYD|Sum-of-years-digits depreciation for a given period|cost;salvage;life;per:Period to calculate
TBILLEQ|Bond-equivalent yield of a Treasury bill|settlement;maturity;discount
TBILLPRICE|Price per $100 of face value of a Treasury bill|settlement;maturity;discount
TBILLYIELD|Yield of a Treasury bill|settlement;maturity;pr
VDB|Declining-balance depreciation over any span of periods|cost;salvage;life;start_period;end_period;?factor;?no_switch:TRUE to never switch to straight-line
XIRR|Internal rate of return for cash flows on irregular dates|values;dates:Date of each cash flow;?guess
XNPV|Net present value for cash flows on irregular dates|rate:Annual discount rate;values;dates:Date of each cash flow
YIELD|Yield of a security that pays periodic interest|settlement;maturity;rate:Annual coupon rate;pr;redemption;frequency;?basis
YIELDDISC|Annual yield of a discounted security|settlement;maturity;pr;redemption;?basis
YIELDMAT|Annual yield of a security that pays interest at maturity|settlement;maturity;issue;rate:Interest rate at issue;pr;?basis

#Date & Time
DATE|Builds a date from year, month and day numbers|year;month;day
DATEDIF|Gap between two dates in whole days, months or years|start_date;end_date;unit:"Y", "M", "D", "MD", "YM" or "YD"
DATEVALUE|Turns a date written as text into a date serial number|date_text:Date in text form
DAY|Day of the month (1-31) from a date|serial_number
DAYS|Number of days between two dates|end_date;start_date
DAYS360|Days between two dates counted on a 360-day year|start_date;end_date;?method:FALSE for the US method, TRUE for European
EDATE|Date that falls a set number of months before or after a date|start_date;months
EOMONTH|Last day of the month a set number of months away|start_date;months
HOUR|Hour (0-23) from a time value|serial_number
ISOWEEKNUM|ISO week number of the year for a date|date:The date to check
MINUTE|Minute (0-59) from a time value|serial_number
MONTH|Month (1-12) from a date|serial_number
NETWORKDAYS|Counts working days between two dates, skipping weekends and holidays|start_date;end_date;?holidays
NETWORKDAYS.INTL|Counts working days between two dates with a custom weekend|start_date;end_date;?weekend;?holidays
NOW|The current date and time, updated on recalculation|
SECOND|Second (0-59) from a time value|serial_number
TIME|Builds a time value from hours, minutes and seconds|hour;minute;second
TIMEVALUE|Turns a time written as text into a time value|time_text:Time in text form
TODAY|The current date, updated on recalculation|
WEEKDAY|Day of the week for a date, as a number|serial_number;?return_type:1 = Sunday-first (default), 2 = Monday-first, and more
WEEKNUM|Week number of the year for a date|serial_number;?return_type:Which weekday starts a week
WORKDAY|Date that is a set number of working days away|start_date;days;?holidays
WORKDAY.INTL|Date a set number of working days away, using a custom weekend|start_date;days;?weekend;?holidays
YEAR|Year from a date|serial_number
YEARFRAC|Portion of a year between two dates|start_date;end_date;?basis

#Engineering
BESSELI|Modified Bessel function In(x)|x;n
BESSELJ|Bessel function Jn(x)|x;n
BESSELK|Modified Bessel function Kn(x)|x;n
BESSELY|Bessel function Yn(x), also called Weber or Neumann function|x;n
BIN2DEC|Converts a binary number to decimal|number:Binary number, up to 10 digits
BIN2HEX|Converts a binary number to hexadecimal|number:Binary number, up to 10 digits;?places
BIN2OCT|Converts a binary number to octal|number:Binary number, up to 10 digits;?places
BITAND|Bitwise AND of two whole numbers|number1:First non-negative integer;number2:Second non-negative integer
BITLSHIFT|Shifts the bits of a number to the left|number:Non-negative integer;shift_amount:Bits to shift (negative shifts right)
BITOR|Bitwise OR of two whole numbers|number1:First non-negative integer;number2:Second non-negative integer
BITRSHIFT|Shifts the bits of a number to the right|number:Non-negative integer;shift_amount:Bits to shift (negative shifts left)
BITXOR|Bitwise exclusive OR of two whole numbers|number1:First non-negative integer;number2:Second non-negative integer
COMPLEX|Builds a complex number from its real and imaginary parts|real_num:Real part;i_num:Imaginary part;?suffix:"i" or "j" (default "i")
CONVERT|Converts a measurement from one unit to another|number:Value to convert;from_unit:Current unit, e.g. "mi";to_unit:Unit for the result, e.g. "km"
DEC2BIN|Converts a decimal number to binary|number:Decimal integer;?places
DEC2HEX|Converts a decimal number to hexadecimal|number:Decimal integer;?places
DEC2OCT|Converts a decimal number to octal|number:Decimal integer;?places
DELTA|Gives 1 when two numbers are equal, otherwise 0|number1:First number;?number2:Second number (default 0)
ERF|Error function integrated between two limits|lower_limit:Lower bound;?upper_limit:Upper bound
ERF.PRECISE|Error function integrated from 0 to x|x
ERFC|Complementary error function from x to infinity|x
ERFC.PRECISE|Complementary error function from x to infinity|x
GESTEP|Gives 1 when a number is at or above a threshold, otherwise 0|number:Value to test;?step:Threshold (default 0)
HEX2BIN|Converts a hexadecimal number to binary|number:Hexadecimal number, up to 10 characters;?places
HEX2DEC|Converts a hexadecimal number to decimal|number:Hexadecimal number, up to 10 characters
HEX2OCT|Converts a hexadecimal number to octal|number:Hexadecimal number, up to 10 characters;?places
IMABS|Absolute value (modulus) of a complex number|inumber
IMAGINARY|Imaginary coefficient of a complex number|inumber
IMARGUMENT|Angle (argument) of a complex number in radians|inumber
IMCONJUGATE|Complex conjugate of a complex number|inumber
IMCOS|Cosine of a complex number|inumber
IMCOSH|Hyperbolic cosine of a complex number|inumber
IMCOT|Cotangent of a complex number|inumber
IMCSC|Cosecant of a complex number|inumber
IMCSCH|Hyperbolic cosecant of a complex number|inumber
IMDIV|Divides one complex number by another|inumber1:Complex numerator;inumber2:Complex denominator
IMEXP|Exponential of a complex number|inumber
IMLN|Natural logarithm of a complex number|inumber
IMLOG10|Base-10 logarithm of a complex number|inumber
IMLOG2|Base-2 logarithm of a complex number|inumber
IMPOWER|Raises a complex number to a power|inumber;number:Power to raise it to
IMPRODUCT|Multiplies complex numbers together|inumber1;inumber2+
IMREAL|Real coefficient of a complex number|inumber
IMSEC|Secant of a complex number|inumber
IMSECH|Hyperbolic secant of a complex number|inumber
IMSIN|Sine of a complex number|inumber
IMSINH|Hyperbolic sine of a complex number|inumber
IMSQRT|Square root of a complex number|inumber
IMSUB|Subtracts one complex number from another|inumber1:Number to subtract from;inumber2:Number to subtract
IMSUM|Adds complex numbers together|inumber1;inumber2+
IMTAN|Tangent of a complex number|inumber
OCT2BIN|Converts an octal number to binary|number:Octal number, up to 10 digits;?places
OCT2DEC|Converts an octal number to decimal|number:Octal number, up to 10 digits
OCT2HEX|Converts an octal number to hexadecimal|number:Octal number, up to 10 digits;?places

#Information
CELL|Reports a detail about a cell, such as its address, format or contents|info_type:Which detail to report, e.g. "address" or "width";?reference:Cell to inspect (default: the formula cell)
ERROR.TYPE|Number that identifies which error a value is|error_val:The error value to identify
INFO|Reports a detail about the current environment|type_text:Which detail to report, e.g. "osversion"
ISBLANK|TRUE if the cell is empty|value
ISERR|TRUE if the value is any error other than #N/A|value
ISERROR|TRUE if the value is any error|value
ISEVEN|TRUE if the number is even|number
ISFORMULA|TRUE if the referenced cell holds a formula|reference
ISLOGICAL|TRUE if the value is TRUE or FALSE|value
ISNA|TRUE if the value is the #N/A error|value
ISNONTEXT|TRUE if the value is anything other than text|value
ISNUMBER|TRUE if the value is a number|value
ISODD|TRUE if the number is odd|number
ISOMITTED|TRUE if a LAMBDA parameter was left out of the call|argument:The LAMBDA parameter to check
ISREF|TRUE if the value is a reference|value
ISTEXT|TRUE if the value is text|value
N|Turns a value into a number (TRUE is 1, text is 0)|value
NA|Returns the #N/A error, meaning no value is available|
SHEET|Index number of a sheet|?value:Sheet name or reference (default: this sheet)
SHEETS|How many sheets a reference spans, or the workbook holds|?reference
TYPE|Code for the kind of value: 1 number, 2 text, 4 logical, 16 error, 64 array|value

#Logical
AND|TRUE only when every condition is TRUE|logical1;logical2+
BYCOL|Runs a LAMBDA on each column and returns one result per column|array;lambda:LAMBDA that takes one column
BYROW|Runs a LAMBDA on each row and returns one result per row|array;lambda:LAMBDA that takes one row
FALSE|The logical value FALSE|
IF|Picks one result when a condition is TRUE and another when FALSE|logical_test:Condition to check;value_if_true:Result when the condition is TRUE;?value_if_false:Result when it is FALSE
IFERROR|Returns a fallback when a formula results in any error|value:Formula or value to check;value_if_error:Result to use on error
IFNA|Returns a fallback when a formula results in #N/A|value:Formula or value to check;value_if_na:Result to use on #N/A
IFS|Checks conditions in order and returns the result for the first TRUE one|logical_test1:First condition;value_if_true1:Result if the first condition is TRUE;?logical_test2+:Next condition;?value_if_true2+:Result for that condition
LAMBDA|Defines a custom, reusable function from parameters and a formula|?parameter1+:Name of an input the function takes;calculation:Formula to evaluate using the parameters
LET|Names intermediate values so a formula can reuse them|name1:First name to define;name_value1:Value given to name1;calculation_or_name2:Final formula, or another name to define;?name_value2+:Value for the next name;?calculation_or_name3+:Final formula, or yet another name
MAKEARRAY|Builds an array of a given size by running a LAMBDA for each cell|rows;columns;lambda:LAMBDA(row, col) that returns each value
MAP|Runs a LAMBDA on each element of one or more arrays|array1:Array to map over;lambda_or_array2+:More arrays, ending with the LAMBDA to apply
NOT|Flips TRUE to FALSE and FALSE to TRUE|logical
OR|TRUE when at least one condition is TRUE|logical1;logical2+
REDUCE|Folds an array into one value by repeatedly applying a LAMBDA|?initial_value:Starting value of the accumulator;array:Array to fold;lambda:LAMBDA(accumulator, value) giving the new total
SCAN|Applies a LAMBDA across an array and returns every running result|?initial_value:Starting value of the accumulator;array:Array to scan;lambda:LAMBDA(accumulator, value) giving the next result
SWITCH|Compares a value to a list of cases and returns the matching result|expression:Value to compare;value1:First case;result1:Result for the first case;?default_or_value2+:Next case, or a final default;?result2+:Result for that case
TRUE|The logical value TRUE|
XOR|TRUE when an odd number of conditions are TRUE|logical1;logical2+

#Lookup & Reference
ADDRESS|Builds a cell address as text from row and column numbers|row_num;column_num;?abs_num:1-4: absolute, mixed or relative;?a1:TRUE for A1 style, FALSE for R1C1;?sheet_text:Sheet name to prefix
AREAS|Number of separate areas in a reference|reference
CHOOSE|Picks a value from a list by its position|index_num:Position to pick, starting at 1;value1:First option;value2+:More options
CHOOSECOLS|Returns only the chosen columns of an array|array;col_num1:First column to keep;col_num2+:More columns to keep
CHOOSEROWS|Returns only the chosen rows of an array|array;row_num1:First row to keep;row_num2+:More rows to keep
COLUMN|Column number of a reference|?reference:Cell to check (default: the formula cell)
COLUMNS|Number of columns in a range or array|array
DROP|Removes rows or columns from the start or end of an array|array;rows:Rows to remove (negative counts from the end);?columns:Columns to remove (negative counts from the end)
EXPAND|Grows an array to a given size, filling new cells|array;rows:Row count of the result;?columns:Column count of the result;?pad_with
FIELDVALUE|Reads a field from a linked data type value|value:Cell holding the linked data type;field_name:Name of the field to read
FILTER|Keeps only the rows or columns that meet a condition|array;include:TRUE/FALSE array that picks what to keep;?if_empty:Result when nothing matches
FORMULATEXT|Shows the formula in a cell as text|reference
GETPIVOTDATA|Pulls a value out of a PivotTable|data_field:Name of the value field;pivot_table:Any cell inside the PivotTable;?field1+:Field to filter on;?item1+:Item in that field
GROUPBY|Groups rows by key columns and aggregates their values|row_fields:Columns to group by;values:Values to aggregate;function:Aggregation such as SUM or a LAMBDA;?field_headers;?total_depth:Whether to show totals and subtotals;?sort_order:Column index to sort by (negative for descending);?filter_array;?field_relationship:0 hierarchy, 1 table
HLOOKUP|Finds a value in the top row of a table and returns a value below it|lookup_value;table_array;row_index_num:Row of the table to return from;?range_lookup
HSTACK|Joins arrays side by side into one wider array|array1;array2+
IMPORTCSV|Loads the contents of a CSV file as an array|path:Path or URL of the CSV file;?skip_rows:Rows to skip at the top;?take_rows:Rows to load after skipping;?locale:Locale used to read numbers and dates
IMPORTTEXT|Loads a delimited text file as an array|path:Path or URL of the text file;?delimiter;?skip_rows:Rows to skip at the top;?take_rows:Rows to load after skipping;?encoding:Character encoding of the file;?locale:Locale used to read numbers and dates
HYPERLINK|Makes a clickable link to a web page, file or location|link_location:Address or path to open;?friendly_name:Text to show in the cell
IMAGE|Places a picture from a web address into a cell|source:HTTPS URL of the image;?alt_text:Description for accessibility;?sizing:0 fit, 1 fill, 2 original size, 3 custom;?height:Height in pixels (custom size);?width:Width in pixels (custom size)
INDEX|Returns the value at a given row and column of a range|array;row_num;?column_num;?area_num:Which area of a multi-area reference to use
INDIRECT|Turns text into a live reference|ref_text:Address or name as text;?a1:TRUE for A1 style, FALSE for R1C1
LOOKUP|Finds a value in one row or column and returns the matching item from another|lookup_value;lookup_vector:Sorted row or column to search;?result_vector:Row or column to return from
MATCH|Position of a value within a row or column|lookup_value;lookup_array;?match_type:1 next smaller (default), 0 exact, -1 next larger
OFFSET|Reference shifted a number of rows and columns from a start|reference:Starting cell or range;rows:Rows to move down (negative moves up);cols:Columns to move right (negative moves left);?height:Rows in the result;?width:Columns in the result
PIVOTBY|Groups by row and column fields and aggregates values into a cross-table|row_fields:Columns that form row groups;col_fields:Columns that form column groups;values:Values to aggregate;function:Aggregation such as SUM or a LAMBDA;?field_headers;?row_total_depth:Whether row totals and subtotals are shown;?row_sort_order:Column index to sort rows by (negative for descending);?col_total_depth:Whether column totals and subtotals are shown;?col_sort_order:Row index to sort columns by (negative for descending);?filter_array;?relative_to:What PERCENTOF-style totals compare against
ROW|Row number of a reference|?reference:Cell to check (default: the formula cell)
ROWS|Number of rows in a range or array|array
RTD|Reads live data from a real-time data COM server|progID:Registered ID of the RTD server;server:Machine running the server (blank for local);topic1:First topic for the data;topic2+:More topics
SORT|Sorts the contents of a range or array|array;?sort_index:Row or column number to sort by;?sort_order;?by_col:TRUE sorts by column instead of row
SORTBY|Sorts an array using the values of other arrays|array;by_array1:Array to sort by;?sort_order1:1 ascending, -1 descending;?by_array2+:Next array to sort by;?sort_order2+:Order for that array
TAKE|Keeps rows or columns from the start or end of an array|array;rows:Rows to keep (negative takes from the end);?columns:Columns to keep (negative takes from the end)
TOCOL|Flattens an array into a single column|array;?ignore:0 keep all, 1 skip blanks, 2 skip errors, 3 skip both;?scan_by_column:TRUE reads column by column
TOROW|Flattens an array into a single row|array;?ignore:0 keep all, 1 skip blanks, 2 skip errors, 3 skip both;?scan_by_column:TRUE reads column by column
TRANSPOSE|Flips rows into columns and columns into rows|array
TRIMRANGE|Cuts empty rows and columns off the edges of a range|range:Range to trim;?trim_rows:0 none, 1 leading, 2 trailing, 3 both (default);?trim_cols:0 none, 1 leading, 2 trailing, 3 both (default)
UNIQUE|Lists the distinct values or rows of a range|array;?by_col:TRUE compares columns instead of rows;?exactly_once:TRUE keeps only items that appear once
VLOOKUP|Finds a value in the first column of a table and returns a value from that row|lookup_value;table_array;col_index_num:Column of the table to return from;?range_lookup
VSTACK|Stacks arrays on top of each other into one taller array|array1;array2+
WRAPCOLS|Wraps a single row or column into columns of a set length|vector:Single row or column to wrap;wrap_count:Values per column;?pad_with
WRAPROWS|Wraps a single row or column into rows of a set length|vector:Single row or column to wrap;wrap_count:Values per row;?pad_with
XLOOKUP|Searches a range and returns the matching item from another range|lookup_value;lookup_array;return_array;?if_not_found;?match_mode;?search_mode
XMATCH|Position of a value in a range, with flexible matching and direction|lookup_value;lookup_array;?match_mode;?search_mode

#Math & Trig
ABS|Absolute value: the number without its sign|number
ACOS|Arccosine of a number, in radians|number:Cosine value between -1 and 1
ACOSH|Inverse hyperbolic cosine of a number|number:Value of 1 or more
ACOT|Arccotangent of a number, in radians|number
ACOTH|Inverse hyperbolic cotangent of a number|number:Value whose magnitude exceeds 1
AGGREGATE|Runs one of 19 summary calculations with options to skip errors or hidden rows|function_num:1-19, e.g. 9 for SUM, 14 for LARGE;options:What to ignore (0-7);ref1:First range or array;ref2+:More ranges, or k for ranking functions
ARABIC|Converts a Roman numeral to a number|text:Roman numeral, e.g. "MMXXVI"
ASIN|Arcsine of a number, in radians|number:Sine value between -1 and 1
ASINH|Inverse hyperbolic sine of a number|number
ATAN|Arctangent of a number, in radians|number
ATAN2|Arctangent of an x/y point, in radians|x_num:X coordinate;y_num:Y coordinate
ATANH|Inverse hyperbolic tangent of a number|number:Value strictly between -1 and 1
BASE|Writes a number in another base as text|number:Non-negative integer;radix:Target base from 2 to 36;?min_length:Pad to at least this many characters
CEILING.MATH|Rounds a number up to the nearest integer or multiple|number;?significance;?mode:Nonzero rounds negatives away from zero
CEILING.PRECISE|Rounds a number up to a multiple, regardless of sign|number;?significance
COMBIN|Number of ways to choose items without repetition|number:Total items;number_chosen:Items in each group
COMBINA|Number of ways to choose items when repeats are allowed|number:Total items;number_chosen:Items in each group
COS|Cosine of an angle|number:Angle in radians
COSH|Hyperbolic cosine of a number|number
COT|Cotangent of an angle|number:Angle in radians
COTH|Hyperbolic cotangent of a number|number
CSC|Cosecant of an angle|number:Angle in radians
CSCH|Hyperbolic cosecant of a number|number
DECIMAL|Converts text in a given base to a decimal number|text:Digits in the source base;radix:Source base from 2 to 36
DEGREES|Converts radians to degrees|angle
EVEN|Rounds away from zero to the nearest even integer|number
EXP|e raised to a power|number:Exponent to raise e to
FACT|Factorial of a number|number:Non-negative number
FACTDOUBLE|Double factorial of a number|number:Non-negative number
FLOOR.MATH|Rounds a number down to the nearest integer or multiple|number;?significance;?mode:Nonzero rounds negatives toward zero
FLOOR.PRECISE|Rounds a number down to a multiple, regardless of sign|number;?significance
GCD|Greatest common divisor of whole numbers|number1;number2+
INT|Rounds a number down to the nearest integer|number
ISO.CEILING|Rounds a number up to a multiple, regardless of sign|number;?significance
LCM|Least common multiple of whole numbers|number1;number2+
LN|Natural logarithm of a number|number:Positive number
LOG|Logarithm of a number to a chosen base|number:Positive number;?base:Base of the logarithm (default 10)
LOG10|Base-10 logarithm of a number|number:Positive number
MDETERM|Determinant of a square matrix|array:Square array of numbers
MINVERSE|Inverse of a square matrix|array:Square array of numbers
MMULT|Matrix product of two arrays|array1:First matrix;array2:Second matrix (rows must equal first's columns)
MOD|Remainder after division|number:Number to divide;divisor:Number to divide by
MROUND|Rounds a number to the nearest multiple|number;multiple:Multiple to round to
MULTINOMIAL|Ratio of the factorial of a sum to the product of factorials|number1;number2+
MUNIT|Identity matrix of a given size|dimension:Number of rows and columns
ODD|Rounds away from zero to the nearest odd integer|number
PERCENTOF|Share that a subset contributes to a total, as a fraction|data_subset:Values in the subset;data_all:All values that make up the total
PI|The constant pi, about 3.14159|
POWER|A number raised to a power|number:Base;power:Exponent
PRODUCT|Multiplies all the given numbers|number1;number2+
QUOTIENT|Whole-number part of a division|numerator:Number to divide;denominator:Number to divide by
RADIANS|Converts degrees to radians|angle:Angle in degrees
RAND|Random number from 0 up to 1, new on each recalculation|
RANDARRAY|Array of random numbers|?rows;?columns;?min:Smallest value (default 0);?max:Largest value (default 1);?whole_number:TRUE for integers
RANDBETWEEN|Random whole number between two limits|bottom;top
ROMAN|Converts a number to Roman numerals as text|number:Value from 1 to 3999;?form:0 classic up to 4 for most concise
ROUND|Rounds a number to a set number of digits|number;num_digits
ROUNDDOWN|Rounds a number toward zero|number;num_digits
ROUNDUP|Rounds a number away from zero|number;num_digits
SEC|Secant of an angle|number:Angle in radians
SECH|Hyperbolic secant of a number|number
SEQUENCE|Array of evenly stepped numbers|rows;?columns;?start:First number (default 1);?step:Increment (default 1)
SERIESSUM|Sum of a power series|x:Input value;n:First power of x;m:Step added to the power each term;coefficients:Coefficient for each term
SIGN|Sign of a number: 1, 0 or -1|number
SIN|Sine of an angle|number:Angle in radians
SINH|Hyperbolic sine of a number|number
SQRT|Positive square root|number:Non-negative number
SQRTPI|Square root of a number multiplied by pi|number:Non-negative number
SUBTOTAL|Summary of a range that skips other subtotals and optionally hidden rows|function_num:1-11 include hidden rows, 101-111 skip them;ref1;ref2+
SUM|Adds numbers together|number1;number2+
SUMIF|Adds cells that meet one condition|range;criteria;?sum_range:Cells to add (default: range)
SUMIFS|Adds cells that meet every one of several conditions|sum_range:Cells to add;criteria_range1;criteria1;?criteria_range2+;?criteria2+
SUMPRODUCT|Multiplies matching items of arrays and adds the results|array1;array2+
SUMSQ|Sum of the squares of the numbers|number1;number2+
SUMX2MY2|Sum of the differences of squares of matching values|array_x:First array;array_y:Second array
SUMX2PY2|Sum of the sums of squares of matching values|array_x:First array;array_y:Second array
SUMXMY2|Sum of the squared differences of matching values|array_x:First array;array_y:Second array
TAN|Tangent of an angle|number:Angle in radians
TANH|Hyperbolic tangent of a number|number
TRUNC|Cuts off the fractional part of a number|number;?num_digits:Digits to keep (default 0)

#Statistical
AVEDEV|Average distance of values from their mean|number1;number2+
AVERAGE|Arithmetic mean of numbers|number1;number2+
AVERAGEA|Mean that also counts text as 0 and TRUE as 1|value1;value2+
AVERAGEIF|Mean of cells that meet one condition|range;criteria;?average_range:Cells to average (default: range)
AVERAGEIFS|Mean of cells that meet every one of several conditions|average_range:Cells to average;criteria_range1;criteria1;?criteria_range2+;?criteria2+
BETA.DIST|Beta distribution|x;alpha;beta;cumulative;?A;?B
BETA.INV|Inverse of the cumulative beta distribution|probability;alpha;beta;?A;?B
BINOM.DIST|Binomial distribution probability|number_s;trials;probability_s;cumulative
BINOM.DIST.RANGE|Chance that successes fall within a range in a binomial trial|trials;probability_s;number_s:Lowest number of successes;?number_s2:Highest number of successes
BINOM.INV|Smallest success count whose cumulative binomial chance reaches a target|trials;probability_s;alpha:Target probability
CHISQ.DIST|Left-tailed chi-squared distribution|x;deg_freedom;cumulative
CHISQ.DIST.RT|Right-tailed chi-squared probability|x;deg_freedom
CHISQ.INV|Inverse of the left-tailed chi-squared distribution|probability;deg_freedom
CHISQ.INV.RT|Inverse of the right-tailed chi-squared distribution|probability;deg_freedom
CHISQ.TEST|Chi-squared test of independence p-value|actual_range:Observed counts;expected_range:Expected counts
CONFIDENCE.NORM|Half-width of a confidence interval using the normal distribution|alpha:Significance level, e.g. 0.05;standard_dev:Population standard deviation;size
CONFIDENCE.T|Half-width of a confidence interval using Student's t distribution|alpha:Significance level, e.g. 0.05;standard_dev:Sample standard deviation;size
CORREL|Correlation coefficient of two data sets|array1;array2
COUNT|Counts cells that hold numbers|value1;value2+
COUNTA|Counts cells that are not empty|value1;value2+
COUNTBLANK|Counts empty cells in a range|range:Range to check
COUNTIF|Counts cells that meet one condition|range;criteria
COUNTIFS|Counts rows that meet every one of several conditions|criteria_range1;criteria1;?criteria_range2+;?criteria2+
COVARIANCE.P|Population covariance of two data sets|array1;array2
COVARIANCE.S|Sample covariance of two data sets|array1;array2
DEVSQ|Sum of squared deviations from the mean|number1;number2+
EXPON.DIST|Exponential distribution|x;lambda:Rate parameter;cumulative
F.DIST|Left-tailed F distribution|x;deg_freedom1;deg_freedom2;cumulative
F.DIST.RT|Right-tailed F probability|x;deg_freedom1;deg_freedom2
F.INV|Inverse of the left-tailed F distribution|probability;deg_freedom1;deg_freedom2
F.INV.RT|Inverse of the right-tailed F distribution|probability;deg_freedom1;deg_freedom2
F.TEST|F-test p-value comparing the variances of two data sets|array1;array2
FISHER|Fisher transformation of a correlation value|x:Value between -1 and 1
FISHERINV|Undoes the Fisher transformation|y
FORECAST.ETS|Forecasts a future value with exponential smoothing (AAA)|target_date;values:Historical values;timeline;?seasonality;?data_completion;?aggregation
FORECAST.ETS.CONFINT|Confidence half-width for an exponential smoothing forecast|target_date;values:Historical values;timeline;?confidence_level:Confidence between 0 and 1 (default 0.95);?seasonality;?data_completion;?aggregation
FORECAST.ETS.SEASONALITY|Season length detected in a time series|values:Historical values;timeline;?data_completion;?aggregation
FORECAST.ETS.STAT|A statistic from the exponential smoothing model|values:Historical values;timeline;statistic_type:Which statistic, 1 to 8;?seasonality;?data_completion;?aggregation
FORECAST.LINEAR|Predicts a value along a straight-line trend|x:Point to predict for;known_y's;known_x's
FREQUENCY|Counts how many values fall into each bin|data_array:Values to count;bins_array:Upper limits of each bin
GAMMA|Gamma function value|x
GAMMA.DIST|Gamma distribution|x;alpha;beta;cumulative
GAMMA.INV|Inverse of the cumulative gamma distribution|probability;alpha;beta
GAMMALN|Natural log of the gamma function|x
GAMMALN.PRECISE|Natural log of the gamma function|x
GAUSS|Chance a standard normal value falls between the mean and z standard deviations|z
GEOMEAN|Geometric mean of positive numbers|number1;number2+
GROWTH|Projects values along an exponential trend|known_y's;?known_x's;?new_x's;?const
HARMEAN|Harmonic mean of positive numbers|number1;number2+
HYPGEOM.DIST|Hypergeometric distribution|sample_s;number_sample;population_s;number_pop;cumulative
INTERCEPT|Where the best-fit line crosses the y axis|known_y's;known_x's
KURT|Kurtosis: how heavy the tails of a data set are|number1;number2+
LARGE|The k-th largest value in a data set|array;k:Rank from the top, 1 = largest
LINEST|Statistics for a least-squares straight-line fit|known_y's;?known_x's;?const;?stats
LOGEST|Statistics for an exponential curve fit|known_y's;?known_x's;?const;?stats
LOGNORM.DIST|Lognormal distribution|x;mean:Mean of ln(x);standard_dev:Standard deviation of ln(x);cumulative
LOGNORM.INV|Inverse of the cumulative lognormal distribution|probability;mean:Mean of ln(x);standard_dev:Standard deviation of ln(x)
MAX|Largest number in a set|number1;number2+
MAXA|Largest value, also counting text as 0 and TRUE as 1|value1;value2+
MAXIFS|Largest value among cells that meet every condition|max_range:Cells to take the maximum of;criteria_range1;criteria1;?criteria_range2+;?criteria2+
MEDIAN|Middle value of a set of numbers|number1;number2+
MIN|Smallest number in a set|number1;number2+
MINA|Smallest value, also counting text as 0 and TRUE as 1|value1;value2+
MINIFS|Smallest value among cells that meet every condition|min_range:Cells to take the minimum of;criteria_range1;criteria1;?criteria_range2+;?criteria2+
MODE.MULT|Vertical array of the most frequent values|number1;number2+
MODE.SNGL|The most frequent value in a data set|number1;number2+
NEGBINOM.DIST|Negative binomial distribution|number_f;number_s:Target number of successes;probability_s;cumulative
NORM.DIST|Normal distribution for a given mean and standard deviation|x;mean;standard_dev;cumulative
NORM.INV|Inverse of the cumulative normal distribution|probability;mean;standard_dev
NORM.S.DIST|Standard normal distribution (mean 0, standard deviation 1)|z;cumulative
NORM.S.INV|Inverse of the cumulative standard normal distribution|probability
PEARSON|Pearson correlation coefficient|array1:Independent values;array2:Dependent values
PERCENTILE.EXC|Value at a percentile, with k strictly between 0 and 1|array;k:Percentile, exclusive of 0 and 1
PERCENTILE.INC|Value at a percentile, with k from 0 to 1|array;k:Percentile from 0 to 1
PERCENTRANK.EXC|Rank of a value as a percentage, excluding 0 and 1|array;x:Value to rank;?significance:Significant digits (default 3)
PERCENTRANK.INC|Rank of a value as a percentage from 0 to 1|array;x:Value to rank;?significance:Significant digits (default 3)
PERMUT|Number of ordered arrangements of a chosen size|number:Total items;number_chosen:Items in each arrangement
PERMUTATIONA|Ordered arrangements when items may repeat|number:Total items;number_chosen:Items in each arrangement
PHI|Density of the standard normal distribution|x
POISSON.DIST|Poisson distribution|x:Number of events;mean:Expected number of events;cumulative
PROB|Chance that values fall between two limits|x_range:Possible values;prob_range:Chance of each value;lower_limit:Lowest value to include;?upper_limit:Highest value to include
QUARTILE.EXC|Quartile of a data set, using exclusive percentiles|array;quart:Which quartile: 1, 2 or 3
QUARTILE.INC|Quartile of a data set, using inclusive percentiles|array;quart
RANK.AVG|Rank of a number in a list, averaging ties|number:Value to rank;ref:List of numbers to rank within;?order:0 descending (default), nonzero ascending
RANK.EQ|Rank of a number in a list, ties share the top rank|number:Value to rank;ref:List of numbers to rank within;?order:0 descending (default), nonzero ascending
RSQ|Square of the Pearson correlation (R squared)|known_y's;known_x's
SKEW|Sample skewness: how lopsided a distribution is|number1;number2+
SKEW.P|Population skewness of a distribution|number1;number2+
SLOPE|Slope of the best-fit straight line|known_y's;known_x's
SMALL|The k-th smallest value in a data set|array;k:Rank from the bottom, 1 = smallest
STANDARDIZE|Converts a value into a z-score|x;mean;standard_dev
STDEV.P|Standard deviation of a whole population|number1;number2+
STDEV.S|Standard deviation estimated from a sample|number1;number2+
STDEVA|Sample standard deviation, counting text as 0 and TRUE as 1|value1;value2+
STDEVPA|Population standard deviation, counting text as 0 and TRUE as 1|value1;value2+
STEYX|Standard error of predicted y values in a regression|known_y's;known_x's
T.DIST|Left-tailed Student's t distribution|x;deg_freedom;cumulative
T.DIST.2T|Two-tailed Student's t probability|x;deg_freedom
T.DIST.RT|Right-tailed Student's t probability|x;deg_freedom
T.INV|Left-tailed inverse of Student's t distribution|probability;deg_freedom
T.INV.2T|Two-tailed inverse of Student's t distribution|probability;deg_freedom
T.TEST|Student's t-test p-value for two data sets|array1;array2;tails;type:1 paired, 2 equal variance, 3 unequal variance
TREND|Projects values along a straight-line trend|known_y's;?known_x's;?new_x's;?const
TRIMMEAN|Mean after dropping a share of the highest and lowest values|array;percent:Fraction of points to drop in total
VAR.P|Variance of a whole population|number1;number2+
VAR.S|Variance estimated from a sample|number1;number2+
VARA|Sample variance, counting text as 0 and TRUE as 1|value1;value2+
VARPA|Population variance, counting text as 0 and TRUE as 1|value1;value2+
WEIBULL.DIST|Weibull distribution|x;alpha;beta;cumulative
Z.TEST|One-tailed p-value of a z-test|array;x:Value to test;?sigma

#Database
DAVERAGE|Average of a column over records that match the criteria|${DB_ARGS}
DCOUNT|Counts numeric cells in a column over matching records|${DB_ARGS}
DCOUNTA|Counts non-empty cells in a column over matching records|${DB_ARGS}
DGET|The single value in a column from the one matching record|${DB_ARGS}
DMAX|Largest value in a column over matching records|${DB_ARGS}
DMIN|Smallest value in a column over matching records|${DB_ARGS}
DPRODUCT|Product of a column over matching records|${DB_ARGS}
DSTDEV|Sample standard deviation of a column over matching records|${DB_ARGS}
DSTDEVP|Population standard deviation of a column over matching records|${DB_ARGS}
DSUM|Total of a column over matching records|${DB_ARGS}
DVAR|Sample variance of a column over matching records|${DB_ARGS}
DVARP|Population variance of a column over matching records|${DB_ARGS}

#Text
ARRAYTOTEXT|Turns an array into a single text string|array;?format:0 concise (default), 1 strict with quotes and braces
ASC|Converts full-width characters to half-width|text
BAHTTEXT|Spells a number out in Thai words with baht|number
CHAR|Character for a code number in the system character set|number:Code from 1 to 255
CLEAN|Removes non-printing characters from text|text
CODE|Code number of the first character of text|text
CONCAT|Joins text from several items or ranges|text1;text2+
COPILOT|Asks the built-in AI assistant a question and returns its answer|prompt_part1:Instruction or context for the AI;?context1+:Cell or range to give the AI as context;?prompt_part2+:More instruction text
DBCS|Converts half-width characters to full-width|text
JIS|Converts half-width characters to full-width (same as DBCS)|text
DETECTLANGUAGE|Identifies the language of a piece of text|text
DOLLAR|Formats a number as currency text|number;?decimals:Digits after the point (default 2)
EXACT|TRUE if two texts are identical, including case|text1:First text;text2:Second text
FIND|Position of one text inside another, case-sensitive|find_text;within_text;?start_num
FINDB|Byte position of one text inside another, case-sensitive|find_text;within_text;?start_num
FIXED|Formats a number as text with fixed decimals|number;?decimals;?no_commas:TRUE leaves out thousands separators
LEFT|Characters from the start of text|text;?num_chars
LEFTB|Bytes from the start of text|text;?num_bytes
LEN|Number of characters in text|text
LENB|Number of bytes in text|text
LOWER|Converts text to lowercase|text
MID|Characters from the middle of text|text;start_num;num_chars
MIDB|Bytes from the middle of text|text;start_num;num_bytes
NUMBERVALUE|Turns text into a number using chosen separators|text;?decimal_separator:Character used as the decimal point;?group_separator:Character used between thousands
PHONETIC|Phonetic (furigana) characters from Japanese text|reference
PROPER|Capitalizes the first letter of each word|text
REGEXEXTRACT|Pulls out the parts of text that match a regular expression|text;pattern;?return_mode:0 first match, 1 all matches, 2 capture groups;?case_sensitivity
REGEXREPLACE|Replaces parts of text that match a regular expression|text;pattern;replacement:Text to put in, may use $1-style groups;?occurrence:Which match to replace (0 = all);?case_sensitivity
REGEXTEST|TRUE if any part of the text matches a regular expression|text;pattern;?case_sensitivity
REPLACE|Swaps a span of characters for new text|old_text;start_num;num_chars:Characters to replace;new_text
REPLACEB|Swaps a span of bytes for new text|old_text;start_num;num_bytes:Bytes to replace;new_text
REPT|Repeats text a given number of times|text;number_times:How many times to repeat
RIGHT|Characters from the end of text|text;?num_chars
RIGHTB|Bytes from the end of text|text;?num_bytes
SEARCH|Position of one text inside another, ignoring case, wildcards allowed|find_text;within_text;?start_num
SEARCHB|Byte position of one text inside another, ignoring case|find_text;within_text;?start_num
SUBSTITUTE|Replaces occurrences of some text with new text|text;old_text:Text to find;new_text;?instance_num
T|Returns the value if it is text, otherwise empty text|value
TEXT|Formats a value as text using a number format code|value;format_text:Format code, e.g. "0.00" or "yyyy-mm-dd"
TEXTAFTER|Text that comes after a delimiter|text;delimiter;?instance_num:Which delimiter to use (negative counts from the end);?match_mode:0 case-sensitive, 1 ignore case;?match_end:1 treats the end of text as a delimiter;?if_not_found
TEXTBEFORE|Text that comes before a delimiter|text;delimiter;?instance_num:Which delimiter to use (negative counts from the end);?match_mode:0 case-sensitive, 1 ignore case;?match_end:1 treats the end of text as a delimiter;?if_not_found
TEXTJOIN|Joins text with a delimiter between items|delimiter:Text to place between items;ignore_empty:TRUE skips empty cells;text1;text2+
TEXTSPLIT|Splits text into columns and rows at delimiters|text;col_delimiter:Where to split into columns;?row_delimiter:Where to split into rows;?ignore_empty:TRUE drops empty pieces;?match_mode:0 case-sensitive, 1 ignore case;?pad_with
TRANSLATE|Translates text from one language to another|text;?source_language:Language code of the text (auto-detect if omitted);?target_language:Language code to translate into
TRIM|Removes extra spaces, leaving single spaces between words|text
UNICHAR|Unicode character for a code point|number:Unicode code point
UNICODE|Code point of the first character of text|text
UPPER|Converts text to uppercase|text
VALUE|Turns text that looks like a number into a number|text
VALUETOTEXT|Turns any value into text|value;?format:0 concise (default), 1 strict with quotes

#Web
ENCODEURL|Percent-encodes text so it is safe inside a URL|text
FILTERXML|Pulls data out of XML text using an XPath expression|xml:XML text to read;xpath:XPath query selecting the data
WEBSERVICE|Fetches the response of a web address as text|url:Address of the web service

#Cube
CUBEKPIMEMBER|A key performance indicator property from the cube|connection;kpi_name:Name of the KPI;kpi_property:Which part: value, goal, status, trend, and so on;?caption
CUBEMEMBER|A member or tuple from the cube, if it exists|connection;member_expression;?caption
CUBEMEMBERPROPERTY|A property value of a cube member|connection;member_expression;property:Name of the property
CUBERANKEDMEMBER|The member at a given rank within a cube set|connection;set_expression;rank:Position to return;?caption
CUBESET|Defines a set of members or tuples on the server|connection;set_expression;?caption;?sort_order:0-6, how to sort the set;?sort_by:Value to sort by
CUBESETCOUNT|Number of items in a cube set|set:Cell holding a CUBESET result
CUBEVALUE|An aggregated value from the cube|connection;?member_expression1+:Member or tuple that filters the value

#Compatibility
BETADIST|Cumulative beta distribution (older version of BETA.DIST)|x;alpha;beta;?A;?B
BETAINV|Inverse beta distribution (older version of BETA.INV)|probability;alpha;beta;?A;?B
BINOMDIST|Binomial distribution (older version of BINOM.DIST)|number_s;trials;probability_s;cumulative
CEILING|Rounds a number up to a multiple (older version of CEILING.MATH)|number;significance
ECMA.CEILING|Rounds a number up to a multiple (same as CEILING)|number;significance
CHIDIST|Right-tailed chi-squared probability (older version of CHISQ.DIST.RT)|x;deg_freedom
CHIINV|Inverse right-tailed chi-squared (older version of CHISQ.INV.RT)|probability;deg_freedom
CHITEST|Chi-squared independence test (older version of CHISQ.TEST)|actual_range:Observed counts;expected_range:Expected counts
CONCATENATE|Joins several text items into one (older version of CONCAT)|text1;text2+
CONFIDENCE|Confidence interval half-width (older version of CONFIDENCE.NORM)|alpha:Significance level, e.g. 0.05;standard_dev:Population standard deviation;size
CALL|Calls a procedure in a code library or resource (legacy add-in feature)|register_id:ID from REGISTER.ID, or a library and procedure name;?argument1+:Values to pass to the procedure
COVAR|Population covariance (older version of COVARIANCE.P)|array1;array2
CRITBINOM|Smallest count reaching a binomial chance (older version of BINOM.INV)|trials;probability_s;alpha:Target probability
EUROCONVERT|Converts between euro and former euro-zone currencies (legacy add-in)|number:Amount to convert;source:ISO code of the currency to convert from;target:ISO code of the currency to convert to;?full_precision:TRUE keeps every digit of the result;?triangulation_precision:Digits used for the intermediate euro value
EXPONDIST|Exponential distribution (older version of EXPON.DIST)|x;lambda:Rate parameter;cumulative
FDIST|Right-tailed F probability (older version of F.DIST.RT)|x;deg_freedom1;deg_freedom2
FINV|Inverse right-tailed F (older version of F.INV.RT)|probability;deg_freedom1;deg_freedom2
FLOOR|Rounds a number down toward zero to a multiple (older version of FLOOR.MATH)|number;significance
FORECAST|Predicts a value along a linear trend (older version of FORECAST.LINEAR)|x:Point to predict for;known_y's;known_x's
FTEST|F-test result (older version of F.TEST)|array1;array2
GAMMADIST|Gamma distribution (older version of GAMMA.DIST)|x;alpha;beta;cumulative
GAMMAINV|Inverse gamma distribution (older version of GAMMA.INV)|probability;alpha;beta
HYPGEOMDIST|Hypergeometric probability (older version of HYPGEOM.DIST)|sample_s;number_sample;population_s;number_pop
LOGINV|Inverse lognormal distribution (older version of LOGNORM.INV)|probability;mean:Mean of ln(x);standard_dev:Standard deviation of ln(x)
LOGNORMDIST|Cumulative lognormal distribution (older version of LOGNORM.DIST)|x;mean:Mean of ln(x);standard_dev:Standard deviation of ln(x)
MODE|Most frequent value (older version of MODE.SNGL)|number1;number2+
NEGBINOMDIST|Negative binomial probability (older version of NEGBINOM.DIST)|number_f;number_s:Target number of successes;probability_s
NORMDIST|Normal distribution (older version of NORM.DIST)|x;mean;standard_dev;cumulative
NORMINV|Inverse normal distribution (older version of NORM.INV)|probability;mean;standard_dev
NORMSDIST|Cumulative standard normal (older version of NORM.S.DIST)|z
NORMSINV|Inverse standard normal (older version of NORM.S.INV)|probability
PERCENTILE|Value at a percentile (older version of PERCENTILE.INC)|array;k:Percentile from 0 to 1
PERCENTRANK|Percentage rank of a value (older version of PERCENTRANK.INC)|array;x:Value to rank;?significance:Significant digits (default 3)
POISSON|Poisson distribution (older version of POISSON.DIST)|x:Number of events;mean:Expected number of events;cumulative
QUARTILE|Quartile of a data set (older version of QUARTILE.INC)|array;quart
RANK|Rank of a number in a list (older version of RANK.EQ)|number:Value to rank;ref:List of numbers to rank within;?order:0 descending (default), nonzero ascending
REGISTER.ID|Registration ID of a code library procedure (legacy add-in feature)|module_text:Name of the library file;procedure:Name of the procedure;?type_text:Data types of the return value and arguments
STDEV|Sample standard deviation (older version of STDEV.S)|number1;number2+
STDEVP|Population standard deviation (older version of STDEV.P)|number1;number2+
TDIST|Student's t probability (older version of T.DIST.RT / T.DIST.2T)|x;deg_freedom;tails
TINV|Two-tailed inverse t (older version of T.INV.2T)|probability;deg_freedom
TTEST|Student's t-test result (older version of T.TEST)|array1;array2;tails;type:1 paired, 2 equal variance, 3 unequal variance
VAR|Sample variance (older version of VAR.S)|number1;number2+
VARP|Population variance (older version of VAR.P)|number1;number2+
WEIBULL|Weibull distribution (older version of WEIBULL.DIST)|x;alpha;beta;cumulative
ZTEST|One-tailed z-test p-value (older version of Z.TEST)|array;x:Value to test;?sigma
`;

const CATEGORIES: ReadonlySet<string> = new Set<FnCategory>([
  'Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database',
  'Text', 'Logical', 'Information', 'Engineering', 'Cube', 'Web', 'Compatibility',
]);

const parseArg = (spec: string, fnName: string): FnArg => {
  const colon = spec.indexOf(':');
  let head = (colon < 0 ? spec : spec.slice(0, colon)).trim();
  const explicit = colon < 0 ? '' : spec.slice(colon + 1).trim();
  const optional = head.startsWith('?');
  if (optional) head = head.slice(1);
  const repeat = head.endsWith('+');
  if (repeat) head = head.slice(0, -1);
  const about = explicit || COMMON_ARGS[head];
  if (!head || !about) throw new Error(`catalog: bad argument "${spec}" in ${fnName}`);
  const arg: FnArg = { name: head, about };
  if (optional || repeat) arg.optional = true;
  if (repeat) arg.repeat = true;
  return arg;
};

const parseCatalog = (src: string): FnInfo[] => {
  const out: FnInfo[] = [];
  let category: FnCategory | null = null;
  for (const raw of src.split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    if (line.startsWith('#')) {
      const c = line.slice(1).trim();
      if (!CATEGORIES.has(c)) throw new Error(`catalog: unknown category "${c}"`);
      category = c as FnCategory;
      continue;
    }
    const [name, about, args = ''] = line.split('|');
    if (!category || !name || !about) throw new Error(`catalog: malformed line "${line}"`);
    out.push({
      name: name.trim(),
      category,
      about: about.trim(),
      args: args.trim() ? args.split(';').map((s) => parseArg(s, name)) : [],
    });
  }
  return out;
};

const freeze = (f: FnInfo): FnInfo => {
  f.args.forEach((a) => Object.freeze(a));
  Object.freeze(f.args);
  return Object.freeze(f);
};

export const CATALOG: readonly FnInfo[] = Object.freeze(
  parseCatalog(DATA).map(freeze).sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0)),
);

const byName = new Map<string, FnInfo>(CATALOG.map((f) => [f.name, f]));

const normalizeName = (name: string): string =>
  name.trim().replace(/^=/, '').replace(/\($/, '').toUpperCase().replace(/^(_XLFN\.|_XLWS\.)+/, '');

export const catalogByName = (name: string): FnInfo | undefined => byName.get(normalizeName(name));

export const searchCatalog = (query: string, limit = 50): FnInfo[] => {
  const q = normalizeName(query);
  const max = Math.max(0, Math.floor(limit));
  if (!q) return CATALOG.slice(0, max);
  const prefix: FnInfo[] = [];
  const contains: FnInfo[] = [];
  const described: FnInfo[] = [];
  const lower = q.toLowerCase();
  for (const f of CATALOG) {
    if (f.name.startsWith(q)) prefix.push(f);
    else if (f.name.includes(q)) contains.push(f);
    else if (f.about.toLowerCase().includes(lower)) described.push(f);
  }
  // Shorter names first among prefix hits so "SUM" beats "SUMIFS" for "su".
  prefix.sort((a, b) => a.name.length - b.name.length || (a.name < b.name ? -1 : 1));
  return [...prefix, ...contains, ...described].slice(0, max);
};

export const signature = (f: FnInfo): string => {
  const parts: string[] = [];
  const { args } = f;
  for (let i = 0; i < args.length; ) {
    const a = args[i];
    if (a.repeat) {
      const group: string[] = [];
      while (i < args.length && args[i].repeat) group.push(args[i++].name);
      parts.push(`[${group.join(', ')}]`, '...');
      continue;
    }
    parts.push(a.optional ? `[${a.name}]` : a.name);
    i++;
  }
  return `${f.name}(${parts.join(', ')})`;
};
