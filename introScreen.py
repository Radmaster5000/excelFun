def intro ():

	print("""

This is a program to take a list of postcodes and retrieve the eastings and northings for each.

When prompted, please enter:

	The full name of the workbook, including extension.

	The name of the worksheet that contains the postcode.

	CURRENTLY:

	The program requires Column A to be the postcode.

	Eastings will be printed in Column B, on the relevant postcode's row.

	Northings will be printed in Column C, on the relevant postcode's row.

Type 's' to start or 'n' for next.

	""")

	answer = input()

	if (answer == 's'):
		return answer
	elif (answer == 'n'):
		print("""

For this to work, please check you have the latest version of geckodriver:

https://github.com/mozilla/geckodriver/releases

Firefox web browser.

and the Selenium webdriver library for Python



	""")
	else:
		print('Not an option')

