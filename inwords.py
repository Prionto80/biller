def number_to_words(num):

    ones = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
            "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen",
            "sixteen", "seventeen", "eighteen", "nineteen"]

    tens = ["", "", "twenty", "thirty", "forty", "fifty",
            "sixty", "seventy", "eighty", "ninety"]

    def two_digit(n):
        if n < 20:
            return ones[n]
        return tens[n // 10] + (" " + ones[n % 10] if n % 10 != 0 else "")

    def three_digit(n):
        if n < 100:
            return two_digit(n)
        return ones[n // 100] + " hundred" + (" " + two_digit(n % 100) if n % 100 != 0 else "")

    if num == 0:
        return "zero"

    parts = []
    
    crore = num // 10000000
    if crore:
        parts.append(two_digit(crore) + " crore")
    num %= 10000000

    lakh = num // 100000
    if lakh:
        parts.append(two_digit(lakh) + " lakh")
    num %= 100000

    thousand = num // 1000
    if thousand:
        parts.append(two_digit(thousand) + " thousand")
    num %= 1000

    if num:
        parts.append(three_digit(num))

    final = " ".join(parts)
    return final.title()
