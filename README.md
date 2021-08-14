# Keyword-Search-in-Specific-Document
The purpose of this project is keyword search in EASA biweekly reports. Keywords include Tc Holder (Manufacturer) and Type (Aircraft type). The program takes data from the .html file and searches for keywords and outputs a processed .xlsx file. If the published instruction is irrelevant for the aircraft manufacturer, "N/A by TC Holder", if it is irrelevant for the aircraft type, "N/A by Type" comments are entered in the excel file. By the way, the instructions published as APPLIANCES are processed automatically as they include every type of aircraft.

To illustrate an example, let's take the 14th biweekly report published on 2021-06-28 as an introduction.

The prepared GUI is shown below.

![app](https://user-images.githubusercontent.com/82766641/129459926-57481eea-b78e-4274-bde8-65baba284c87.png)

In the application, you need to select a file with .html extension, otherwise the program will crash and close.
You need to enter Keywords like this way: Words, Searching.

![APP1](https://user-images.githubusercontent.com/82766641/129460033-8ddf1885-2b1e-4f2b-bf38-27142971e3c8.png)

For example, keywords are entered for Boeing 737 aircraft type.
## Input

![İNPUT](https://user-images.githubusercontent.com/82766641/129460212-9705fa6b-8001-44a4-b610-4ed53c6bbf47.png)


## Output


![OUT](https://user-images.githubusercontent.com/82766641/129460189-debbf9c5-ab9a-4d94-8b86-a2b6dff559dc.png)
