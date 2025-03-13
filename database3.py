import sqlite3
import customtkinter as ctk
import pandas as pd
import openpyxl 
from tkinter import messagebox

class PhotoStudioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Photo Studio Management")
        self.root.geometry("600x400")
        
        #Database Connection
        self.conn = sqlite3.connect("photo_studio.db")
        self.cursor = self.conn.cursor()
        self.create_tables()
        
        self.create_tabs()
    
    def create_tables(self):
        """Creates the necessary database tables if they do not exist."""
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS bookings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_name TEXT,
                booking_date TEXT,
                amount REAL
            )
        """)  # Corrected SQL syntax
        self.conn.commit()
    
    def create_tabs(self):
        """Creates the GUI elements."""
        self.frame = ctk.CTkFrame(self.root)
        self.frame.pack(pady=20)
        
        #  Customer Name
        self.label_name = ctk.CTkLabel(self.frame, text="Customer Name:")
        self.label_name.grid(row=0, column=0, padx=10, pady=5)
        self.entry_name = ctk.CTkEntry(self.frame)
        self.entry_name.grid(row=0, column=1, padx=10, pady=5)
        
        # Booking Date
        self.label_date = ctk.CTkLabel(self.frame, text="Booking Date:")
        self.label_date.grid(row=1, column=0, padx=10, pady=5)
        self.entry_date = ctk.CTkEntry(self.frame)
        self.entry_date.grid(row=1, column=1, padx=10, pady=5)
        
        #  Amount
        self.label_amount = ctk.CTkLabel(self.frame, text="Amount:")
        self.label_amount.grid(row=2, column=0, padx=10, pady=5)
        self.entry_amount = ctk.CTkEntry(self.frame)
        self.entry_amount.grid(row=2, column=1, padx=10, pady=5)
        
        #  Add Booking Button
        self.button_add = ctk.CTkButton(self.frame, text="Add Booking", command=self.add_booking)
        self.button_add.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Export to Excel Button
        self.button_export = ctk.CTkButton(self.frame, text="Export to Excel", command=self.export_to_excel)
        self.button_export.grid(row=4, column=0, columnspan=2, pady=10)
        
        #  Booking List
        self.booking_list = ctk.CTkTextbox(self.root, width=500, height=150)
        self.booking_list.pack(pady=10)
        
        self.load_bookings()
    
    def add_booking(self):
        """Adds a new booking to the database."""
        name = self.entry_name.get()
        date = self.entry_date.get()
        amount = self.entry_amount.get()
        
        if name and date and amount:
            try:
                self.cursor.execute("INSERT INTO bookings (customer_name, booking_date, amount) VALUES (?, ?, ?)", (name, date, float(amount)))
                self.conn.commit()
                self.load_bookings()
                messagebox.showinfo("Success", "Booking added successfully!")
            except Exception as e:
                messagebox.showerror("Database Error", f"Error: {e}")
        else:
            messagebox.showerror("Error", "Please fill all fields")
    
    def load_bookings(self):
        """Loads all bookings from the database and displays them in the UI."""
        self.booking_list.delete("1.0", "end")
        self.cursor.execute("SELECT customer_name, booking_date, amount FROM bookings")
        bookings = self.cursor.fetchall()
        if bookings:
            for row in bookings:
                self.booking_list.insert("end", f"{row[0]} - {row[1]} - ₹{row[2]}\n")
        else:
            self.booking_list.insert("end", "No bookings found.")
    
    def export_to_excel(self):
        """Exports the database data to an Excel file."""
        self.cursor.execute("SELECT * FROM bookings")
        data = self.cursor.fetchall()
        if data:
            df = pd.DataFrame(data, columns=["ID", "Customer Name", "Booking Date", "Amount"])
            try:
                df.to_excel("PhotoStudioData.xlsx", index=False, engine="openpyxl")  # ✅ FIXED: Using openpyxl engine
                messagebox.showinfo("Success", "Data exported to Excel successfully!")
            except Exception as e:
                messagebox.showerror("Excel Export Error", f"Error: {e}")
        else:
            messagebox.showerror("Error", "No data to export")
    
    def __del__(self):
        """Closes the database connection when the application exits."""
        self.conn.close()

if __name__ == "__main__":
    root = ctk.CTk()
    app = PhotoStudioApp(root)
    root.mainloop()
