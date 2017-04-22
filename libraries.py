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
redFill = PatternFill('solid',fgColor='FF0000')
greenFill = PatternFill('solid',fgColor='00FF00')
yellowFill = PatternFill('solid',fgColor='FFFF00')

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
    print('result: ', query_result)
    if len(query_result.places)== 0:
        not_found_libraries.append(library_name)
        return 'not found'
    else:
        pass
    for place in query_result.places:
        print('place: ', place.name)
        place.get_details()
        try:
            print(place.details['permanently_closed'])
            closed_libraries.append(place.name)
            return False
        except KeyError:
            open_libraries.append(place.name)
            return True



"""
The loop below loops through values in column F (col 6)
and runs the above function using that value. Depending on the result
of the function run, the cell color is changed.
"""
for i in range(2,50): # change '40' to '4041' to run script on entire column
    col = 6
    library = (ws.cell(row=i,column=col).value)
    if checkIfClosed(library, i)== True:
        ws.cell(row=i,column=col).fill=greenFill
    elif checkIfClosed(library, i) == False:
        ws.cell(row=i,column=col).fill=redFill
    elif checkIfClosed(library, i) == 'not found':
        ws.cell(row=i, column=col).fill = yellowFill

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


