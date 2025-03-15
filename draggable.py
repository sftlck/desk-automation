import tkinter as tk
import pyodbc
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Draggable Text Boxes")
        self.root.geometry("500x300")

        # Connect to the MS Access database
        self.db_connection = self.connect_to_db()

        # Initialize the next available ID only if the database connection is successful
        if self.db_connection:
            self.next_id = self.get_next_card_id()
        else:
            self.next_id = 1  # Fallback if the database connection fails

        # Add Activity Button
        self.add_button = tk.Button(root, text="Add Activity", command=self.add_text_box)
        self.add_button.pack(side="top", pady=10)

        # Counter for text box labels
        self.counter = 1

    def connect_to_db(self):
        # Define the connection string for MS Access
        db_path = r"C:\Users\Castro\Desktop\Computa\Draggable\database-d.accdb"  # Update this path
        if not os.path.exists(db_path):
            print(f"Error: Database file not found at {db_path}")
            return None

        connection_string = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + db_path
        try:
            connection = pyodbc.connect(connection_string)
            print("Connected to the database successfully!")
            return connection
        except Exception as e:
            print(f"Error connecting to the database: {e}")
            return None

    def get_next_card_id(self):
        if not self.db_connection:
            print("Database connection is not available.")

        cursor = self.db_connection.cursor()
        query = "SELECT MAX(id) FROM activities"
        cursor.execute(query)
        result = cursor.fetchone()
        cursor.close()

        if result[0] is not None:
            return result[0] + 1 
        else:
            return 1 
    def add_text_box(self):
        if self.db_connection:
            text = f"Activity {self.next_id}"
            DraggableTextBox(self.root, text, 50, 50, self.db_connection, self)

            self.next_id += 1
        else:
            print("Database connection failed. Cannot add activity.")

class DraggableTextBox:
    def __init__(self, master, text, x, y, db_connection, app):
        self.master = master
        self.text = text
        self.db_connection = db_connection
        self.app = app
        self.label = tk.Label(master, text=text, bg="white", relief="raised", padx=10, pady=5)
        self.label.place(x=x, y=y)

        self.label.bind("<Button-1>", self.on_drag_start)
        self.label.bind("<B1-Motion>", self.on_drag_motion)
        self.label.bind("<ButtonRelease-1>", self.on_drag_release)
        self.label.bind("<Double-Button-1>", self.on_double_click) 

    def on_drag_start(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def on_drag_motion(self, event):
        x = self.label.winfo_x() - self.start_x + event.x
        y = self.label.winfo_y() - self.start_y + event.y
        self.label.place(x=x, y=y)

    def on_drag_release(self, event):
        x, y = self.label.winfo_x(), self.label.winfo_y()
        snapped_x, snapped_y = self.snap_to_position(x, y)
        self.label.place(x=snapped_x, y=snapped_y)

        self.save_coordinates_to_db(snapped_x, snapped_y)

    def on_double_click(self, event):
        self.open_edit_window()

    def open_edit_window(self):
        self.edit_window = tk.Toplevel(self.master)
        self.edit_window.title("Edit Activity")
        self.edit_window.geometry("400x200")
        self.edit_window.attributes("-topmost", True)
        self.edit_window.grab_set()

        tk.Label(self.edit_window, text="Activity Name:").pack(pady=10)
        self.activity_name_entry = tk.Entry(self.edit_window, width=30)
        self.activity_name_entry.pack(pady=5)
        self.activity_name_entry.insert(0, self.text)

        tk.Label(self.edit_window, text="Technician:").pack(pady=10)
        self.activity_name_entry2 = tk.Entry(self.edit_window, width=30)
        self.activity_name_entry2.pack(pady=5)
        self.activity_name_entry2.insert(0, "")

        save_button = tk.Button(self.edit_window, text="Save", command=self.save_activity_name)
        save_button.pack(pady=10)

    def save_activity_name(self):
        new_name = self.activity_name_entry.get()
        self.text = new_name
        self.label.config(text=new_name)
        self.save_name_to_db(new_name)
        self.edit_window.destroy()

    def save_name_to_db(self, new_name):
        if self.db_connection:
            cursor = self.db_connection.cursor()
            query = "UPDATE activities SET activity_name = ? WHERE id = ?"
            cursor.execute(query, (new_name, self.app.next_id - 1))  # Use the correct ID
            self.db_connection.commit()
            cursor.close()
        else:
            print("Database connection failed. Cannot save activity name.")

    def snap_to_position(self, x, y):
        # Define the 5 predefined positions
        positions = [
            (50, 50),   # Position 1
            (200, 50),  # Position 2
            (350, 50),  # Position 3
            (50, 150),  # Position 4
            (200, 150), # Position 5
        ]

        # Find the nearest position
        min_distance = float('inf')
        snapped_pos = (x, y)
        for pos in positions:
            distance = (x - pos[0]) ** 2 + (y - pos[1]) ** 2
            if distance < min_distance:
                min_distance = distance
                snapped_pos = pos
        return snapped_pos

    def save_coordinates_to_db(self, x, y):
        if self.db_connection:
            cursor = self.db_connection.cursor()
            query = "INSERT INTO activities (activity_name, x_coord, y_coord) VALUES (?, ?, ?)"
            cursor.execute(query, (self.text, x, y))
            self.db_connection.commit()
            cursor.close()
        else:
            print("Database connection failed. Cannot save coordinates.")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()