# Old way of timing (Outdated)


# Number of values that the average will be calculated from. The more
# values used, the more accurate results will be. Adjust this value if
# runtime is too long or too short
precision = 100

# Make a list of timings to get a more accurate mean
timings = [ timeit.timeit('solve', 'from __main__ import solve', number=1000) for x in range(precision + 1) ]

# The first time is always much longer, so delete this one
del timings[0]

# Remove the timings that are 0, and replace them with other timings
timings_filtered = [ ]
for timing in timings:
	if timing != 0:
		timings_filtered.append(timing)
	else:
		x = 0
		while x == 0:
			x = timeit.timeit('solve', 'from __main__ import solve', number=1000)

		timings_filtered.append(x)

# Convert the average time to scientific notation for readability
average = sum(timings_filtered)/precision
average = '%.4E' % Decimal(average)

# Calculate the mean time from the list of values
print("Average time:", average, "sec")