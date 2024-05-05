# Populate Lat Long Data in Excel Using Python

This _**site_lat_ling.py**_ script reads an Excel file containing location information (street, city, state and zip) and fetch the latitude, longitude and formatted address using the Google Geocoding API for each location in the spreadsheet.

# Resources

* [Python 3](https://www.python.org/doc/)
* [Google Geocoding API](https://developers.google.com/maps/documentation/geocoding/overview)
* [Creation of Virtual Environments](https://packaging.python.org/en/latest/guides/installing-using-pip-and-virtual-environments/#create-and-use-virtual-environments)

# Run the program

Requirements
* You have Python 3.10 installed on your computer
* You created a virtual environment in the project directory.
* You activated the virtual environment.
* Execute `pip install -r requirements.txt` to install dependencies.
* Use the provided `sites.xlsx` file as a template for your locations.
* Create a .env file in the project directory and provide your the following variables:

    * `GOOGLE_MAP_API_KEY=<Insert your Google key here>`
    * `ENV_PATH='<Insert path to .env file here>'`

* Create an empty directory named `logs` in the project directory.

* To run script, type `python3 sites_lat_long.py`, then click `ENTER`