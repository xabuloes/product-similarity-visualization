#!/usr/bin/python3
# -*- coding: utf-8 -*-

#imports
from xlrd import open_workbook #to process the xlsx file
import sys #to process the given arguments

#filepath to the product reviews stored in Excel
global file_path
file_path = 'Data_NMD.xlsx'

'''
filepath to the file containing words that are too common to be keywords
in german or make no sense
'''
global german_default_file
german_default_file = "german_default_file.txt"

'''
filepath to the file containing words that are too common to be keywords
in english or make no sense
'''
global english_default_file
english_default_file = "english_default_file.txt"

#special characters, that should be escaped from the reviews
global special_characters
special_characters = "!\"§$%&/()=?´`²³[]}{\\*+~'#_-:.;,µ><|°^ß@€♥•"

'''
global variable that manages if the frequency of the keywords in the reviews
is printed or not.
Do not change the value of the variable here, but change the related variable
"word_count" in the main function that calls the corresponding setter functions.
The value 1 indicates, that the frequency of the keywords is printed.
The value 0 indicates, that the frequency of the keywords is NOT printed.
'''
global show_word_count
show_word_count = 1

'''
global variable that manages the number of top keywords that are printed.
Do not change the value of the variable here, but change the related variable
"xx" in the main function that calls the corresponding setter function.
'''
global top
top = 10

'''
global variable that manages the language of wich the the reviews are processed.
Do not change the value of the variable here, it is changed by the parameters
given by the programm caller.
The value 0 indicates, that only ENGLISH reviews are used.
The value 1 indicates, that only GERMAN reviews are used.
'''
global language
language = 0

'''
prints a final information, about how many different products the xlsx file
does contain and how many different reviews exist

the parameter different_products is a list of all different products existing
in the xlsx file and created by the function get_different_products

the parameter values must be a list of all product reviews as created by the
function read_excel_file() or an already filtered list from the functions
filter_for_english_entries/filter_for_german_entries
'''
def final_info_print(different_products,values):

    #fetch current setted language
    global language

    #print corresponding information indicators
    if(language):
        dp = "\nVerschiedene Produkte: "
        de = "Verschiedene Einträge: "
    else:
        dp = "\ndifferent products: "
        de = "different entries: "

    #print information
    #number of different products
    print(dp + str(len(different_products)))
    #number of different reviews
    print(de + str(len(values)))

'''
escapes default language words that can not be considered as key words

the parameter dictionary is a list of the words used in the reviews sorted
by their frequency
this parameter is provided by the function sorted_by_word_frequency
'''
def escape_default_words(dictionary):

    #access escape word files
    global german_default_file
    global english_default_file

    #check out which language is currently selected and open the correct file
    global language
    if(language):
        file = open(german_default_file)
    else:
        file = open(english_default_file)

    #create a empty list for the default words
    default_words = list()

    #read the file word by word and append them to the created list
    line = file.readline()
    while(line != ''):
        default_words.append(str(line)[:-1])
        line = file.readline()
    file.close()
    
    #iterate over default_words and delete them from dictionary
    for i in range(len(default_words)):
        try:
            dictionary.remove(default_words[i])
        except:
            pass

    #return the list of remaining words
    return dictionary

'''
creates a list of the words used in the reviews sorted by their frequency and
escape default language words that can not be considered key words

the parameter dictionary is the dictionary containing the words in the reviews
with the corresponding occuring frequency
'''
def sorted_by_word_frequency(dictionary):

    #sort the dictionary and write it to a list
    dictionary = sorted(dictionary, key=dictionary.__getitem__)
    dictionary = dictionary[::-1]

    #escape default language words that can not be considered key words
    dictionary = escape_default_words(dictionary)

    #return the list
    return dictionary

'''
gets all reviews of a product merged in the string merged_product_reviews 
created by the function merge_product_review and returns a dictionary of all
words with their occuring frequency in the string
'''
def word_count(merged_product_reviews):
    
    #create empty dictionary
    counts = dict()

    #split the merged_product_reviews by whitespace into a list of words
    words = merged_product_reviews.split()

    #iterate over the whole word list and do for each word in words
    for word in words:

        #if the dictionary already contains the word
        if word in counts:
            #increment the corresponding dictionary entry,
            #functioning as word count, by 1
            counts[word] += 1
        else:
            #else insert the word to the dictionary with a word count 1
            counts[word] = 1

    #return the dictionary with words and corresponding frequencies
    return counts

'''
escapes all special characters of the given review string created within
the function merge_product_review
'''
def escape_special_characters(reviews):

    #access the characters that should be escaped
    global special_characters

    #iterate over whole review
    for i in range(len(special_characters)):

        #replace special characters throug a whitespace
        reviews = reviews.replace(special_characters[i],' ')

    #return the new review string
    return reviews

'''
merges all review texts to one huge text

The parameter product is a one entry of the list of different_products.
The list different_products is a list of all different products existing
in the xlsx file and created by the function get_different_products

The parameter values must be a list of all product reviews as created by the
function read_excel_file() or an already filtered list from the functions
filter_for_english_entries/filter_for_german_entries.
'''
def merge_product_review(product, values):

    #create empty review string
    reviews = ""

    #iterate over all reviews
    for i in range(len(values)):
        #if review refers to product concatinate the review with reviews
        if(product == values[i][0:3]):
            reviews += " " + values[i][4] + " " + values[i][5]

    #change all letters to lower case for easy word comparison
    reviews = reviews.lower()

    #escape all special characters of the review string
    reviews = escape_special_characters(reviews)

    #return the review string
    return reviews

'''
creates a product name string out of the value of product

The parameter product is a one entry of the list of different_products.
The list different_products is a list of all different products existing
in the xlsx file and created by the function get_different_products
'''
def create_product_name(product):

    return product[0] + " " + product[1] + " " + product[2]

'''
prints the product name and its corresponding top keywords

The parameter product is a one entry of the list of different_products.
The list different_products is a list of all different products existing
in the xlsx file and created by the function get_different_products

The parameter values must be a list of all product reviews as created by the
function read_excel_file() or an already filtered list from the functions
filter_for_english_entries/filter_for_german_entries.
'''
def map_most_used_words(product, values):

    #TODO print to files???

    #create a product name string out of the value of product
    product_name = create_product_name(product)

    #merge all review texts to one huge text
    merged_product_reviews = merge_product_review(product, values)

    '''
    create a dictionary of the words in the merged reviews an their
    corresponding frequency
    '''
    dictionary_word_ocurrance = word_count(merged_product_reviews)

    '''
    create a list of the words used in the reviews sorted by their frequency and
    escape default language words that can not be considered key words
    '''
    word_occurance = sorted_by_word_frequency(dictionary_word_ocurrance)

    #select wich language is choosen and set product string
    if(language):
        product_string = "Produkt "
    else:
        product_string = "Product "

    #print an empty line for beauty adjustments
    print()
    #print the product string and the product name
    print(product_string + str(product_name) + ":")
    #print an empty line for beauty adjustments
    print()

    #lookup how many top keywords should be printed
    global top

    '''
    iterate over the word_occurance list and print the top values with the
    corresponding frequency from dictionary_word_ocurrance if requested

    all this is done in a try block because sometimes there do not remain top
    words
    '''
    for i in range(top):

        #check if frequency should be shown
        if(show_word_count):
            try:
                #print word plus frequency
                print(str(word_occurance[i]) + ":" + str(dictionary_word_ocurrance.get(str(word_occurance[i]))))
            except:
                pass
        else:
            try:
                #just print word
                print(str(word_occurance[i]))
            except:
                pass

    #print an empty line for beauty adjustments
    print()

'''
creates a list of all different products

the parameter values is a list of all product reviews as created by the function
read_excel_file() or an already filtered list from the functions
filter_for_english_entries/filter_for_german_entries
'''
def get_different_products(values):

    #create empty list
    different_products = list()

    #iterate over values
    for i in range(len(values)):

        #if the product does not exists in different products, add it
        if(not(values[i][0:3] in different_products)):
            different_products.append(values[i][0:3])

    #return the list of different_products
    return different_products

'''
filters the received list values for all entries from english speaking countries
and does not consider all entries from another language

the parameter values is a list of all product reviews as created by the function
read_excel_file()
'''
def filter_for_english_entries(values):

    #create empty list for the filtered values
    english_values = list()

    '''
    iterate over all values and append the reviews with the language tag
    containing "en" to the list of filtered entries
    '''
    for i in range(len(values)):
        if("en" in values[i][3]):
           english_values.append(values[i])

    #return list of only english entries
    return english_values

'''
filters the received list values for all entries from german speaking countries
and does not consider all entries from another language

the parameter values is a list of all product reviews as created by the function
read_excel_file()
'''
def filter_for_german_entries(values):

    #create empty list for the filtered values
    german_values = list()

    '''
    iterate over all values and append the reviews with the language tag
    containing "de" to the list of filtered entries
    '''
    for i in range(len(values)):
        if("de" in values[i][3]):
           german_values.append(values[i])

    #return list of only german entries
    return german_values

'''
reads the content of the xlsx File given by the variable file_path.
return a list of all the entries in the File.
An entry in the list consists of a list containig each row of the xlsx File
divided by its columns.

The parameter file_path contains the file_path for the xlsx file to read
'''
def read_excel_file(file_path):

    #open file from file_path
    wb = open_workbook(file_path)

    #iterate over file an create the returning list
    for s in wb.sheets():
        values = []
        for row in range(s.nrows):
            col_value = []
            for col in range(s.ncols):
                value  = (s.cell(row,col).value)
                try : value = str(int(value))
                except : pass
                col_value.append(value)
            values.append(col_value)

    return values

#maps and prints the keywords related to each product
def map_and_print_keywords():

    '''
    take global file_path where the product review xlsx file is stored and
    write all its content into a list
    '''
    global file_path
    original_values = read_excel_file(file_path)

    #checkout wich language is selected
    global language
    if(language):
        #filter the received list original_values for all german entries
        values = filter_for_german_entries(original_values)
    else:
        #filter the received list original_values for all english entries
        values = filter_for_english_entries(original_values)

    #creates a list of all different products
    different_products = get_different_products(values)

    #iterate over the list of all different products
    for i in range(len(different_products)):
        #print the product name and its corresponding top keywords
        map_most_used_words(different_products[i], values)

    '''
    print a final information, about how many different products the xlsx file
    does contain and how many different reviews exist
    '''
    final_info_print(different_products,values)

#changing the value of the global variable top to the value of xx
def print_top_xx_entries(xx):
    global top
    top = xx

#changes the global variable word_count, that the word count is NOT printed.
def dont_show_word_count():
    global show_word_count
    show_word_count = 0

#changes the global variable word_count, that the word count is printed.
def do_show_word_count():
    global show_word_count
    show_word_count = 1

#setting the global variable language to german
def set_language_german():
    global language
    language = 1

#setting the global varialbe language to english
def set_language_english():
    global language
    language = 0

def arguments_routine():

    '''
    check for enough arguments, print fancy not enough arguments messages and
    exit properly
    '''
    try:
        language_tag = sys.argv[1]
    except IndexError:
        print()
        print('\x1b[0;30;41m' + "Not enough arguments!" + '\x1b[0m')
        print('\x1b[1;33;40m' + \
              "Usage: ./keywords.py --language <language>" + '\x1b[0m')
        print('\x1b[1;33;40m' + "for help use option [--help]" + '\x1b[0m')
        print()
        sys.exit(1)

    '''
    check if the help option was choosen, print the help message and
    exit properly
    '''
    if(language_tag == '-h' or language_tag == '--help' \
       or language_tag == '-help'):
        print()
        print("usage: ./keywords.py [-h] -l <language>")
        print()
        print("arguments:")
        print("-h, --help       show this help message and exit")
        print("-l, --language   set language to english or german")
        print()
        print("<language> needs to be a value of the set:\n[\"en\",\"english\","
               +"\"englisch\",\"de\",\"german\",\"deutsch\"]")
        print()
        sys.exit(1)

    '''
    check if the language option was choosen, if not print invalid arguments
    message and exit properly
    '''
    if(not(language_tag == '-l' or language_tag == '--language')):
        print()
        print ('\x1b[0;30;41m' + "Invalid arguments!" + '\x1b[0m')
        print('\x1b[1;33;40m' + \
              "Usage: ./keywords.py --language <language>" + '\x1b[0m')
        print('\x1b[1;33;40m' + "for help use option [--help]" + '\x1b[0m')
        print()
        sys.exit(1)

    '''
    check if the language option was given with a corresponding parameter,
    if not print not enough arguments message and exit properly
    '''
    try:
        language_value = sys.argv[2]
    except IndexError:
        print()
        print('\x1b[0;30;41m' + "Not enough arguments!" + '\x1b[0m')
        print('\x1b[1;33;40m' + \
              "Usage: ./keywords.py --language <language>" + '\x1b[0m')
        print('\x1b[1;33;40m' + "for help use option [--help]" + '\x1b[0m')
        print()
        sys.exit(1)

    '''
    check wich language was choosen and select the language
    if an unknown language was choosen, print invalid arguments message,
    show wich values language can be and exit properly
    '''
    print(language_value[0:2])
    if(language_value[0:2] == 'en'):
        #setting the language to english
        set_language_english()
    elif(language_value[0:2] == 'de' or language_value[0:6] == 'german'):
        #setting the language to german
        set_language_german()
    else:
        print()
        print ('\x1b[0;30;41m' + "Invalid arguments!" + '\x1b[0m')
        print('\x1b[1;33;40m' + \
              "Usage: ./keywords.py --language <language>" + '\x1b[0m')
        print('\x1b[1;33;40m' + "<language> needs to be a value of the set:" + 
              "\n[\"en\",\"english\",\"englisch\",\"de\",\"german\",\"deutsch" + 
              "\"]" + '\x1b[0m')
        print('\x1b[1;33;40m' + "for help use option [--help]" + '\x1b[0m')
        print()
        sys.exit(1)

def main():

    #check for rigth arguments for the programm
    arguments_routine()

    #value of word_count manages if the frequency of the keywords is printed
    word_count = True
    if(word_count):
        do_show_word_count()
    else:
        dont_show_word_count()

    #value of xx changes the number printed top values
    xx = 10
    print_top_xx_entries(xx)

    #finally map and print the keywords related to each product
    map_and_print_keywords()

if __name__ == "__main__":
    main()