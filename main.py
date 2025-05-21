import pandas as pd
from rapidfuzz import fuzz, process
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, messagebox
import os
import spacy
from PIL import Image, ImageTk
from geopy.geocoders import Nominatim
import folium
import webbrowser

# Load correction data
current_dir = os.path.dirname(os.path.abspath(__file__))
correction_file_path = os.path.join(current_dir, 'Clean_Street_Names.csv')

# Load Swedish spaCy model
nlp = spacy.load("sv_core_news_sm")

# Load corrected data from CSV file
try:
    corrected_data = pd.read_csv(correction_file_path, delimiter=';', on_bad_lines='skip')
    corrected_data['Streetname'] = corrected_data['Streetname'].str.lower()  # Normalize street names to lowercase
    corrected_data['Postalcode'] = corrected_data['Postalcode'].astype(str)  # Ensure postal codes are string type
except Exception as e:
    print(f"Error loading correction data: {e}")
    exit()

# Function to extract street number and optional letter (e.g., "123A")
def extract_street_number(street_name):
    if pd.isna(street_name):  # Check if the street name is missing
        return None

    if ',' in street_name:
        pre_comma_match = re.search(r'\b(\d{1,3}[A-Za-z])(?=,)', street_name)  # Capture number-letter pattern before comma
        if pre_comma_match:
            return pre_comma_match.group(1)

    pattern = r'\b(\d{1,3})(?:-\d{1,3})?(?:\s*([A-Za-z](?!/|,)))?\b'  # Main pattern for number and optional letter
    match = re.search(pattern, street_name)

    if match:
        number = match.group(1)
        letter = match.group(2)

        if number.startswith('0'):  # Exclude leading zero numbers
            return None

        if letter:
            return f"{number} {letter}"  # Combine number and letter if present
        return number  # Only return the number

    return None

# Function to split street names (e.g., "Main St/2nd Ave" -> ["Main St", "2nd Ave"])
def split_street_names(text):
    return [name.strip() for name in re.split(r'[\/, ]+', text) if name.strip()]

# Function to extract street names using Named Entity Recognition (NER)
def extract_street_names_with_ner(text):
    doc = nlp(text)
    street_names = []
    
    # Capture recognized geographical or location entities
    for ent in doc.ents:
        if ent.label_ == "GPE" or ent.label_ == "LOC":
            street_names.append(ent.text)
    
    return street_names

# Define the regex pattern for street names
street_pattern = r'\b([A-Za-z]+(?:\s[A-Za-z]+)*(?:\s*(vägen|gatan|allén|avenyn|boulevard|torg|väg)))\b'

# Function to find best matches for street name and postal code using fuzzy matching
def find_best_matches(street_name, postal_code):
    all_matches = []

    possible_street_names = split_street_names(street_name)
    
    for individual_street_name in possible_street_names:
        street_names = extract_street_names_with_ner(individual_street_name)
        if not street_names:
            street_names = [individual_street_name]
        
        filtered_data = corrected_data[corrected_data['Postalcode'] == postal_code]  # Filter by postal code

        cleaned_street_name = individual_street_name.replace('/', ' ').replace(',', ' ').replace('.', ' ').strip()  # Clean street name for matching

        # Find best matches using fuzzy matching
        name_matches = process.extract(cleaned_street_name.lower(), 
                                       filtered_data['Streetname'].str.lower().tolist(), 
                                       scorer=fuzz.token_set_ratio, 
                                       limit=3)

        match_found = False
        for match in name_matches:
            match_score = fuzz.token_set_ratio(cleaned_street_name.lower(), match[0])
            if match_score >= 90:  # If match score is sufficiently high
                matched_row = filtered_data[filtered_data['Streetname'].str.lower() == match[0]]
                if not matched_row.empty:
                    if abs(len(cleaned_street_name) - len(match[0])) > 3:
                        continue
                    all_matches.append({
                        "Original_Streetname": cleaned_street_name.capitalize(),
                        "Corrected_Streetname": matched_row['Streetname'].iloc[0].capitalize(),
                        "Match_Score": match_score
                    })
                    match_found = True
                    break

        # If no match found, try matching individual words within the street name
        if not match_found:
            for word in street_names:
                if len(word) < 3:
                    continue

                word_matches = process.extract(word.lower(), filtered_data['Streetname'].tolist(), scorer=fuzz.token_set_ratio, limit=3)
                for match in word_matches:
                    match_score = max(
                        fuzz.ratio(word.lower(), match[0]),
                        fuzz.partial_ratio(word.lower(), match[0])
                    )
                    if match_score >= 90:
                        matched_row = filtered_data[filtered_data['Streetname'] == match[0]]
                        if not matched_row.empty:
                            if abs(len(word) - len(match[0])) > 3:
                                continue
                            all_matches.append({
                                "Original_Word": word.capitalize(),
                                "Corrected_Streetname": matched_row['Streetname'].iloc[0].capitalize(),
                                "Match_Score": match_score
                            })
                            match_found = True
                            break
                if match_found:
                    break
                
        if not all_matches:
            pattern_matches = re.findall(street_pattern, cleaned_street_name)
            for match in pattern_matches:
                all_matches.append({
                    "Original_Streetname": match[0].capitalize(),
                    "Corrected_Streetname": match[0].capitalize(),
                    "Match_Score": 100
                })

    return all_matches if all_matches else None  # Return matches or None if no matches were found

# Function to load and process the Excel file
def process_excel_file(file_path, progress_bar, percent_label, cancel_flag, close_button):
    try:
        input_data = pd.read_excel(file_path)

        street_column = None
        postal_column = None

        if 'DeliveryStreet' in input_data.columns:
            street_column = 'DeliveryStreet'
        elif 'Streetname' in input_data.columns:
            street_column = 'Streetname'

        if 'DeliveryZipCode' in input_data.columns:
            postal_column = 'DeliveryZipCode'
        elif 'PostalCode' in input_data.columns:
            postal_column = 'PostalCode'

        if not street_column:
            messagebox.showerror("Error", "The file must contain 'DeliveryStreet' or 'Streetname' column.")
            return
        if not postal_column:
            messagebox.showerror("Error", "The file must contain 'DeliveryZipCode' or 'PostalCode' column.")
            return

        input_data['Corrected_Streetname'] = None
        total_rows = len(input_data)

        for index, row in input_data.iterrows():
            if cancel_flag[0]:
                messagebox.showinfo("Cancelled", "Operation cancelled by the user.")
                return

            street_name = row.get(street_column, 'No Data Provided')
            postal_code = str(row.get(postal_column, ''))

            if pd.isna(street_name):
                input_data.at[index,'Corrected_Streetname'] = "No Data Provided"
                continue

            try:
                result = find_best_matches(street_name, postal_code)
                if result:
                    street_number = extract_street_number(street_name)
                    corrected_streetnames = list(set([match['Corrected_Streetname'] for match in result]))
                    combined_street = '/ '.join(corrected_streetnames)
                    if street_number:
                        combined_street = f"{combined_street} {street_number}"
                    input_data.at[index, 'Corrected_Streetname'] = combined_street
                else:
                    input_data.at[index,'Corrected_Streetname'] = "No Match Found"
            except Exception as e:
                input_data.at[index, 'Corrected_Streetname'] = f"Error: {str(e)}"

            progress = min(int(((index + 1) / total_rows) * 100), 100)
            progress_bar['value'] = progress
            percent_label.config(text=f"{progress}%")
            root.update_idletasks()

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Save Corrected File")
        if save_path:
            input_data.to_excel(save_path, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"File saved successfully at:\n{save_path}")

        close_button.pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to clean the data
def clean_data(progress_bar, percent_label):
    file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        try:
            input_data = pd.read_excel(file_path)
            input_data = input_data.fillna('Unknown')

            for column in input_data.select_dtypes(include=['object']).columns:
                if column not in ['DeliveryZipCode', 'PostalCode']:
                    input_data[column] = input_data[column].str.replace(r'\b[Cc]/[oO]\b', '', regex=True)
                    input_data[column] = input_data[column].apply(
                        lambda x: re.sub(r'[.,/(){}<>!@#$%^&*;:"|?]', ' ', str(x))
                    )

            save_path = file_path.replace(".xlsx", "_Cleaned_Ready_for_Processing.xlsx")
            input_data.to_excel(save_path, index=False, engine='openpyxl')

            messagebox.showinfo("Success", f"The data is now ready for processing. The cleaned file is saved at:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Function to add placeholder text to Entry
def add_placeholder(entry_widget, placeholder_text):
    def on_focus_in(event):
        if entry_widget.get() == placeholder_text:
            entry_widget.delete(0, tk.END)
            entry_widget.config(fg='black')

    def on_focus_out(event):
        if not entry_widget.get():
            entry_widget.insert(0, placeholder_text)
            entry_widget.config(fg='gray')

    entry_widget.insert(0, placeholder_text)
    entry_widget.config(fg='gray')
    entry_widget.bind("<FocusIn>", on_focus_in)
    entry_widget.bind("<FocusOut>", on_focus_out)


# Function to get locality (city) from postal code
def get_locality_by_postal_code(postal_code):
    try:
        result = corrected_data.loc[corrected_data['Postalcode'] == postal_code, 'Locality']
        return result.iloc[0] if not result.empty else "Unknown"
    except Exception as e:
        print("Data error:", e)
        return "Unknown"

# Function for instant search by postal code, street name, and extracted street names
def instant_search():
    postal_code = postal_input.get()
    if postal_code == "Postal Code" or not postal_code.strip():
        postal_code = ""
    
    street_name = street_input.get()
    if not street_name.strip() or street_name == "Street Name":
        messagebox.showwarning("Input Error", "Please enter a street name.")
        return

    street_number = extract_street_number(street_name)

    # Search with postal code first
    result = find_best_matches(postal_code=postal_code, street_name=street_name) if postal_code else []

    # If no results from postal code, search by street name
    if not result:
        result = find_best_matches(street_name=street_name, postal_code=postal_code)

    # If no results from street name, search by extracted street names
    if not result:
        result = find_best_matches(street_name=extract_street_number(street_name), postal_code=postal_code)

    if result:
        unique_results = {res['Corrected_Streetname'] for res in result}

        result_window = tk.Toplevel(root)
        result_window.title("Search Result")
        result_window.geometry("400x500")

        tk.Label(result_window, text="Corrected Street Names:", font=("Arial", 12)).pack(pady=5)

        for result_name in unique_results:
            result_entry = tk.Entry(result_window, font=("Arial", 12), width=40)
            result_entry.insert(0, result_name)
            result_entry.pack(pady=5)

        if street_number:
            tk.Label(result_window, text="Street Number:", font=("Arial", 12)).pack(pady=5)
            number_entry = tk.Entry(result_window, font=("Arial", 12), width=40)
            number_entry.insert(0, street_number)
            number_entry.pack(pady=5)

        if postal_code:
            tk.Label(result_window, text="Postal Code:", font=("Arial", 12)).pack(pady=5)
            postal_entry = tk.Entry(result_window, font=("Arial", 12), width=40)
            postal_entry.insert(0, postal_code)
            postal_entry.pack(pady=5)

            # Fetch and display the locality (city)
            locality = get_locality_by_postal_code(postal_code)
            tk.Label(result_window, text="City:", font=("Arial", 12)).pack(pady=5)
            locality_entry = tk.Entry(result_window, font=("Arial", 12), width=40)
            locality_entry.insert(0, locality)
            locality_entry.pack(pady=5)

            # Create a formatted address string
            formatted_address = f"{', '.join(unique_results)} {street_number}, {locality}" if street_number else f"{', '.join(unique_results)}, {locality}"

            tk.Label(result_window, text="Address:", font=("Arial", 12)).pack(pady=5)
            formatted_entry = tk.Entry(result_window, font=("Arial", 12), width=40)
            formatted_entry.insert(0, formatted_address)
            formatted_entry.pack(pady=5)

        # Add a button to show the map
        map_button = tk.Button(result_window, text="Show on Map", font=("Arial", 12),
                               command=lambda: show_map_for_address(result_name, postal_code))
        map_button.pack(pady=10)

        close_button = tk.Button(result_window, text="Close", font=("Arial", 12), command=result_window.destroy)
        close_button.pack(pady=10)
    else:
        messagebox.showinfo("No Match", "No matches found for the provided postal code or street name.")
        
# Function to get the coordinates (latitude and longitude) for an address
def get_coordinates(address):
    geolocator = Nominatim(user_agent="address_locator")
    location = geolocator.geocode(address)
    if location:
        return location.latitude, location.longitude
    else:
        return None, None

# Function to show the map for a given address
def show_map_for_address(street_name, postal_code):
    locality = get_locality_by_postal_code(postal_code)
    formatted_address = f"{street_name}, {locality}, {postal_code}"

    # Get the coordinates of the address
    latitude, longitude = get_coordinates(formatted_address)

    if latitude and longitude:
        # Create a folium map centered at the address
        map_object = folium.Map(location=[latitude, longitude], zoom_start=15)

        # Add a marker for the address
        folium.Marker([latitude, longitude], popup=formatted_address).add_to(map_object)

        # Save the map as an HTML file
        map_object.save("address_map.html")

        # Open the map in the default web browser
        import webbrowser
        webbrowser.open("address_map.html")
    else:
        messagebox.showerror("Error", "Could not find coordinates for the address.")

# Tkinter UI setup
root = tk.Tk()
root.title("Address Correction")
root.geometry("500x600")

window_width = 500
window_height = 700
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x_pos = (screen_width - window_width) // 2
y_pos = (screen_height - window_height) // 2

root.geometry(f'{window_width}x{window_height}+{x_pos}+{y_pos}')

root.iconbitmap("Logo.ico")

cancel_flag = [False]

try:
    logo = Image.open("Front_logo.png")
    logo = logo.resize((130, 180))
    logo_img = ImageTk.PhotoImage(logo)
    
    logo_label = tk.Label(root, image=logo_img)
    logo_label.pack(pady=10)
except Exception as e:
    messagebox.showerror("Error", f"Failed to load logo: {e}")

tk.Label(root, text="Address Correction Tool", font=("Arial", 16, "bold")).pack(pady=10)

def clear_inputs():
    street_input.delete(0, tk.END)
    postal_input.delete(0, tk.END)
    clear_button.pack_forget()

def on_text_change(event):
    if street_input.get() or postal_input.get():
        clear_button.pack(pady=10)
    else:
        clear_button.pack_forget()

street_frame = tk.Frame(root)
street_frame.pack(pady=5, anchor='w', padx=10)

street_label = tk.Label(street_frame, text="Street Name  ", font=("Arial", 12), fg="black")
street_label.pack(side='left')
street_input = tk.Entry(street_frame, font=("Arial", 12), width=40)
street_input.pack(side='left')
add_placeholder(street_input, "Street Name")

postal_frame = tk.Frame(root)
postal_frame.pack(pady=5, anchor='w', padx=10)

postal_label = tk.Label(postal_frame, text="Postal Code  ", font=("Arial", 12), fg="black")
postal_label.pack(side='left')
postal_input = tk.Entry(postal_frame, font=("Arial", 12), width=40)
postal_input.pack(side='left')
add_placeholder(postal_input, "Postal Code")

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

clear_button = tk.Button(button_frame, text="Clear", font=("Arial", 12), command=clear_inputs)
clear_button.pack(side='left', padx=5)

search_button = tk.Button(button_frame, text="Search", font=("Arial", 12), command=instant_search)
search_button.pack(side='left', padx=5)

street_input.bind("<KeyRelease>", on_text_change)
postal_input.bind("<KeyRelease>", on_text_change)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

percent_label = tk.Label(root, text="0%", font=("Arial", 12))
percent_label.pack()

def open_file_dialog(progress_bar, percent_label, cancel_flag, close_button):
    file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        process_excel_file(file_path, progress_bar, percent_label, cancel_flag, close_button)

clean_button_frame = tk.Frame(root)
clean_button_frame.pack(pady=10)

clean_button = tk.Button(clean_button_frame, text="Clean Data", font=("Arial", 12), command=lambda: clean_data(progress_bar, percent_label))
clean_button.pack(side='left')

description_label = tk.Label(clean_button_frame, text="Step 1: Prepare and Clean Data for Processing", font=("Arial", 10), fg="gray")
description_label.pack(side='left', padx=10)

process_button_frame = tk.Frame(root)
process_button_frame.pack(pady=10)

process_button = tk.Button(process_button_frame, text="Process Excel File", font=("Arial", 12),
                           command=lambda: open_file_dialog(progress_bar, percent_label, cancel_flag, close_button))
process_button.pack(side='left')

description_label = tk.Label(process_button_frame, text="Step 2: Process and Correct Data Errors", font=("Arial", 10), fg="gray")
description_label.pack(side='left', padx=10)

close_button = tk.Button(root, text="Close", font=("Arial", 12), command=root.quit)
close_button.pack(pady=40)

footer_label = tk.Label(root, text="Designed by Ahmad Zalkat", font=("Arial", 8), fg="gray")
footer_label.pack(side="bottom", pady=5)

root.mainloop()