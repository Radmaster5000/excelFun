def checkSpeed ():
	print("""

Please rate the speed of your internet on a scale of 1 - 5.

The connection must be fast enough to load a page in teh estimated time.

If the page does not load in the estimated time, the program may miss an Easting or Northing.

1 = page will load within one second

3 = page will load within three seconds

5 = page will load within five seconds


		""")

	return input('Speed: ')