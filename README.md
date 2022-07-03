# Google-Scholar-Profile-Scrapper
ABSTRACT

The program takes the Google Scholar’s profile’s URL as input and creates an
excel sheet with various details in output. The sheet will contain the various
information scraped from the profile such as Name, Description, Specializations,
various articles written, their description, authors and various citation statistics.

● TOOLS AND TECHNOLOGIES USED

The project is a Node.JS based program which uses various libraries to achieve the
work. The modules used and their description are as follows:
1. “minimist” -> It helps in taking input from the console in an easy way. parse
argument options.
This module is the guts of an optimist's argument parser without all the fanciful
decoration.
2. “axios” -> Promise based HTTP client for the browser and node.js.
Axios provides a simple to use library in a small package with a very extensible
interface.
3. “jsdom” -> jsdom is a pure-JavaScript implementation of many web standards,
notably the WHATWG DOM and HTML Standards, for use with Node.js. In
general, the goal of the project is to emulate enough of a subset of a web
browser to be useful for testing and scraping real-world web applications.
4. “excel4node” -> A full featured xlsx file generation library allowing for the
creation of advanced Excel files.
5. “fs” -> The fs module provides a lot of very useful functionality to access and
interact with the file system.
There is no need to install it. Being part of the Node.js core, it can be used by
simply requiring it.

● SCOPE OF THE PROJECT

Functional Requirement:
The user of the program can be able to get the excel file of the desired
person’s Google Scholar account.

Non-Functional Requirement:
➢ The data representation in the excel file must be organized and
clean.
➢ The download time of the HTML of the webpage must be least.
➢ The Excel creation and update time must be as minimum as
possible.

● DESIGN OF THE PROJECT

The project first of all downloads the HTML of the webpage using axios.
let responsePromise = axios.get(args.source);
The downloaded HTML is then converted to DOM and then Document which then
becomes ready to extract any information from the webpage.

The various informations is then extracted from the webpage mainly using

➢ querySelector()

➢ querySelectorAll()

➢ getElementsByClassName()

➢ getElementsById()

The Excel workBook is then created using ‘excel4node’ and sheets are added with
data imputed, which was scraped from the webpage.
The program even downloads the image of the person using ‘fs’ and ‘request’.
