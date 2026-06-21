function PROPERWORDS(amount) {

  // Convert input to number
  amount = Number(amount);

  // Check invalid values
  if (isNaN(amount) || amount === null || amount === "") {
    return "";
  }

  if (amount === 0) return "Zero Pesos Only";

  var ones = [
    "", "one", "two", "three", "four", "five", "six",
    "seven", "eight", "nine", "ten", "eleven", "twelve",
    "thirteen", "fourteen", "fifteen", "sixteen",
    "seventeen", "eighteen", "nineteen"
  ];

  var tens = [
    "", "", "twenty", "thirty", "forty",
    "fifty", "sixty", "seventy", "eighty", "ninety"
  ];

  function convertHundreds(n) {
    var word = "";

    if (n > 99) {
      word += ones[Math.floor(n / 100)] + " hundred ";
      n %= 100;
    }

    if (n > 19) {
      word += tens[Math.floor(n / 10)] + " ";
      n %= 10;
    }

    if (n > 0) {
      word += ones[n] + " ";
    }

    return word.trim();
  }

  function convert(n) {

    n = Math.floor(n);

    if (n < 20) {
      return ones[n];
    }

    if (n < 1000) {
      return convertHundreds(n);
    }

    if (n < 1000000) {
      return convert(Math.floor(n / 1000)) +
        " thousand " +
        convert(n % 1000);
    }

    if (n < 1000000000) {
      return convert(Math.floor(n / 1000000)) +
        " million " +
        convert(n % 1000000);
    }

    return convert(Math.floor(n / 1000000000)) +
      " billion " +
      convert(n % 1000000000);
  }

  var pesos = Math.floor(amount);
  var centavos = Math.round((amount - pesos) * 100);

  var words = convert(pesos)
    .replace(/\b\w/g, l => l.toUpperCase())
    .trim();

  if (centavos > 0) {

    var centsWords = convert(centavos)
      .replace(/\b\w/g, l => l.toUpperCase())
      .trim();

    return words + " Pesos And " + centsWords + " Centavos Only";
  }

  return words + " Pesos Only";
}