import numpy as np
import openpyxl
import pandas as pd
import requests


def get_book_genre(book_title):

    response = requests.get(
        "https://www.googleapis.com/books/v1/volumes?q={}&key=YOUR_API_KEY".format(
            book_title, "YOUR_API_KEY"))
    authors = ["Unknown"]
    genre = ["Unknown"]
    description = "Unknown"

    if response.status_code == 200:
        book_info = response.json().get("items", [{}])[0]
        volume_info = book_info.get("volumeInfo", {})
        if "authors" in volume_info:
            authors  = volume_info["authors"]
        if "categories" in volume_info:
            genre  = volume_info["categories"]
        if "description" in volume_info:
            description = volume_info["description"]

    # authors_str = ", ".join(authors)
    # genre_str = ", ".join(genre)
   
    #return authors_str, genre_str,  description
    return authors, genre,  description

def classify_book(book_title):
    authors, genre , description = get_book_genre(book_title)
    if "Unknown" in description:
        description = "Description not available"
    if "Unknown" in authors: 
        authors = ["Author not specified"]
    if "Unknown" in genre:
        genre = ["Genre not specified"]

    authors_str = ", ".join(authors)
    genre_str = ", ".join(genre)

    print (authors_str)
    print(genre_str)
    print(description)
    return authors_str , genre_str, description

def main():

    excel_file_path = input("Enter the path to the Excel file: ")

    wb = openpyxl.load_workbook(excel_file_path)

    ws = wb.active
    values = ws.values

    column_names = ['Book Title', 'Authors' , 'Genre', 'Description']

    df = pd.DataFrame(values, columns=column_names)
    df["Authors"], df["Genre"], df["Description"] = zip(*df["Book Title"].apply(classify_book))

    new_excel_file_path = excel_file_path[:-5] + "_classified.xlsx"
    df.to_excel(new_excel_file_path, index=False)

    print("The books have been classified and saved to the following Excel file:")
    print(new_excel_file_path)


if __name__ == "__main__":
    main()
