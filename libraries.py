"""
In order for this code to work, you will need to install the googleplaces and openpyxl libraries. I used pip to install both.
You can find many tutorials on install each online. If you have questions, you can let me know. 

-Openpyxl will not allow the excel file to be open while the program is running or it throw an error. 
-You'll see as the program runs that I have search results printed to the console. I did for a couple reasons, one because
the program takes a while to run and this provides feedback that it's continuing to function and also because you can see 
how many results were returned in the search. 
-Another functionality I want to add is the ability to also log results where more
than 1 result is returned. Currently, I can't be sure that when multiple results are returned, Google is giving me the value I'm
looking for. Their documentation shows that the results are ranked by prominence (a value calculated by Google).

Googleplaces API:
You will need to sign up to use google api services. The first level sign up includes 1000 queries. If you attach
a credit card to the account, you will be permitted 150,000 queries. Google says you won't be charged for this level
and it is just for verification. You can read more about that on the Googleplaces API page.

You will need to edit code on lines 24,27, and 80 with your own information. I've left comments to the side of those lines
exlaining what is needed.
"""

from googleplaces import GooglePlaces, types, lang
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, PatternFill, colors


API_KEY = '' #Enter your key between quotes
google_places = GooglePlaces(API_KEY)

wb = load_workbook('library.xlsx')
ws = wb.active


"""
These variables are used to define fill colors
for changing cell colors based on library search
results
"""
redFill = PatternFill('solid',fgColor='FF0000') #Found one exact match and is closed
greenFill = PatternFill('solid',fgColor='00FF00') #Found one exact match and is open
yellowFill = PatternFill('solid',fgColor='FFFF00') #Did not find any matches
cyanFill = PatternFill('solid',fgColor='00FFFF') #Found multiple results, none exact, is open
magentaFill = PatternFill('solid',fgColor='FF00FF') #Found multiple results, none exact, is closed

"""
Set of list to show all library results at end of script
"""
not_found_libraries = []
closed_libraries = []
open_libraries = []

def checkIfClosed(library_name, row):
    """
    Searches for a library and then finds whether library
    is permanently closed, is open, or is not found.
    Input: library name
    Returns: 'not found': if search returns nothing
            'False': if library is permanently closed
            'True': if library is still open
    """
    query_result = google_places.nearby_search(
        location=str(ws.cell(row=row, column=9).value)+','+str(ws.cell(row=row,column=11).value), keyword=str(library_name),
        rankby='distance')
    print('Number of results found: ', len(query_result.places))
    if len(query_result.places)== 0:
        not_found_libraries.append(library_name)
        print('no results found')
        return 'not found'
    else:
        place, match = findBestResult(query_result.places, library_name)
        place.get_details()
        try:
            print(place.details['permanently_closed'])
            closed_libraries.append(place.name)
            return False, match
        except KeyError:
            open_libraries.append(place.name)
            return True, match

def findBestResult(results, library_name):
    """This function is used to choose best result when
       more than one result is found. If none of the results
       match the library name exactly, the most 'prominent'
       result is returned. 'Prominence' is determined by Google's
       alogrithms.
    """
    for place in results:
        if str(place.name) == str(library_name):
            return place, None
        else:
            print('searched for: ', library_name,)
            print('found instead: ', place.name)
            pass
        print('did not find exact match, returning: ', results[0].name,'\n')
        return results[0], 'notexact'  

"""
The loop below loops through values in column F (col 6)
and runs the above function using that value. Depending on the result
of the function run, the cell color is changed.
"""
for i in range(2,40): # first number (2) tells script to start on row2, second number tells to stop at row 39. Enter (2,4041) to run all rows.
    col = 6
    library = (ws.cell(row=i,column=col).value)
    values = checkIfClosed(library,i)
    if values[0]==True and values[1]==None:
        ws.cell(row=i,column=col).fill=greenFill
    elif values[0] == False and values[1]==None:
        ws.cell(row=i,column=col).fill=redFill
    elif values == 'not found':
        ws.cell(row=i, column=col).fill = yellowFill
    elif values[0] == True and values[1] == 'notexact':
        ws.cell(row=i, column=col).fill = cyanFill
    elif values[0] == False and values[1] == 'notexact':
        ws.cell(row=i, column=col).fill = magentaFill

wb.save('library.xlsx') # saves changes to excel workbook



"""
Following print statements print a list of
libraries which were not found, permanently closed, or
still open. This may be useful for double checking closed
libraries or not found libraries. The values which are printed
in these list reflect the text Google used to search.
"""
print('These libraries were not found in search:')
print(not_found_libraries, '\n')
print('These libraries are permanently closed:')
print(closed_libraries, '\n')
print('These libraries are still open:')
print(open_libraries)
