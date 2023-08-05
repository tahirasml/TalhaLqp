					Data Acquisition (data_acquisition.ipynb)

In this phase, the objective is to retrieve data from the OpenWeatherMap API for a given city (in our case, London) and save it in a CSV format.

Steps:

Import necessary libraries: requests for API calls and pandas for data manipulation.
Define the api_key and target city.
Build the base_url to make the API call.
Send the GET request to the OpenWeatherMap API to fetch the current weather data.
Convert the API response to a JSON format.
Extract the 'main' section from the JSON data, which contains weather details.
Convert the extracted data into a Pandas DataFrame.
Save the DataFrame to a weather_data.csv file for further processing.


					Data Pre-processing (data_preprocessing.ipynb)
In this phase, the raw data acquired from the API is cleaned and transformed to make it suitable for analysis.

Steps:

Load the raw data from weather_data.csv using Pandas.
Check and handle missing values, if any. In our example, we simply drop rows with any missing values.
Convert the temperature data from Kelvin to Celsius.
Save the cleaned and transformed data to preprocessed_weather_data.csv.


					Data Analysis (data_analysis.ipynb)
In this phase, we perform exploratory data analysis (EDA) to understand our dataset's characteristics and derive insights.

Steps:

Load the preprocessed data from preprocessed_weather_data.csv using Pandas.
Calculate basic statistical measures for the temperature:
Mean: Average temperature.
Median: Middle value in the sorted list of temperatures.
Standard Deviation: Measures the amount of variation or dispersion of the temperature set.
Print out the statistical measures.
Visualize the temperature distribution using a histogram. This gives an idea about the frequency distribution of the temperature values.