/**
 * Adds two numbers.
 * @customfunction
 * @param {number} a First number
 * @param {number} b Second number
 * @returns {number} The sum
 */
function add(a, b) {
  return a + b;
}

/**
 * Returns the input value (identity function).
 * @customfunction
 * @param {string} value The value to return
 * @returns {string} The same value
 */
function get(value) {
  return value;
}

CustomFunctions.associate("ADD", add);
CustomFunctions.associate("GET", get);
