# microchip-stock-check
A python script that queries and returns the part availability for a given list of part numbers in an excel file. An update by me for the internal Microchip Technologies tool. 

This tool upon my retrieval only queried microchip, digikey, and mouser. 

# My updates
 * Added Arrow support 
 * Added Avnet Americas support 
 * Added Newark support 
 * Added Future Electronics support
 * Added menu that allows for searching through different MPN search sets.
 * Added ctrl-c instant quit
 * Added filtered search. (reduce a massive list by only searching for specific packages and parts)

# Notes 
* The 'MPN' files are excel spreadsheets that contain part numbers to be searched for. They can be manually lengthened or shortened in excel itself, or can be shortened in program by the filtered search tool. 

* Search sets MUST start with 'MPN' and end in '.xlsx'. You can change the start value by changing the 'startswith' variable. 
