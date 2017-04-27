from googleplaces import GooglePlaces, types, lang
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, PatternFill, colors


API_KEY = 'AIzaSyDLdDH5kqx4-kYy3qlTpxjGc0KEVD9orN4' #Enter your key between quotes
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
blueFill = PatternFill('solid',fgColor='0000FF') #Not exact match found and address of best result does not match address of searched librarys

not_found_libraries = []
closed_dict = {}
open_dict = {}
wrong_address = {}

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
        location=str(ws.cell(row=row, column=9).value)+','+str(ws.cell(row=row,column=11).value),
        keyword=str(library_name)+' '+str(ws.cell(row=row, column=9).value),
        rankby='distance')
    print('Number of results found: ', len(query_result.places))
    print('that result is: ', query_result.places)
    if len(query_result.places)== 0:
        not_found_libraries.append(library_name)
        return 'not found'
    else:
        place, match = findBestResult(query_result.places, library_name)
        place.get_details()
        address = addressChange(place.formatted_address)
        if addressCheck(address, row):
            print('address check: True')
            try:
                _ = (place.details['permanently_closed'])
                closed_dict[libary_name]={'found':place.name, 'address':place.formatted_address}
                return False, match
            except KeyError:
                open_dict[library_name]={'found':place.name, 'address':place.formatted_address}
                return True, match
        else:
            print('address check: False')
            wrong_address[library_name] = {'correct address':str(ws.cell(row=row, column=7).value),
                                           'found address':place.formatted_address}
            return 'no match'

def findBestResult(results, library_name):
    for place in results:
        if str(place.name) == str(library_name):
            return place, None
        else:
            pass
        return results[0], 'notexact'

def addressCheck(address, row):
    """This function removes numbers from address
       and returns True if street matches and False
       if they do not match.
    """
    known_address = str(ws.cell(row=row, column =7).value).strip()
##    print('known address: ', known_address)
##    print('address: ', address)
    return address == known_address

def addressChange(address):
    fixed_address = address[:address.index(',')]
    return fixed_address
    

"""
The loop below loops through values in column F (col 6)
and runs the above function using that value. Depending on the result
of the function run, the cell color is changed.
"""
for i in range(2,20): # change '40' to '4041' to run script on entire column
    col = 6
    library = (ws.cell(row=i,column=col).value)
    values = checkIfClosed(library,i)
    if values == 'no match':
        ws.cell(row=i, column=col).fill=blueFill
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
print('closed dict: ', closed_dict)

print('The following libraries are open')
for k in open_dict:
    print('searched for: ', k)
    print('found: ', open_dict[k]['found'])
    print('address: ', open_dict[k]['address'])

for k in wrong_address:
    print('found: ', k)
    print('found; ', wrong_address[k]['correct address'])
    print('address: ', wrong_address[k]['found address'])


