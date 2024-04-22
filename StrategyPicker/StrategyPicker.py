import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from strategies import strategies
import os
import webbrowser

def find_matching_strategies(price_outlook, volatility_outlook, profit_potential, risk):
    return [strategy for strategy in strategies if
            strategy["price_outlook"] == price_outlook and
            strategy["volatility_outlook"] == volatility_outlook and
            strategy["profit_potential"] == profit_potential and
            strategy["risk"] == risk]

def create_excel_file(strategy_name):
    current_dir = os.path.dirname(__file__)  # Get the current directory of the script
    file_path = os.path.join(current_dir, "strategy_files", f"{strategy_name}.xlsx")  
    actual_file_path = os.path.abspath(file_path)
    if os.path.isfile(actual_file_path):
        print(f"Opening Excel file: {actual_file_path}")
        webbrowser.open(actual_file_path)
    else:
        print(f"Excel file for {strategy_name} does not exist.")


def display_matched_strategies(matched_strategies):
    num_matched_strategies = len(matched_strategies)
    if num_matched_strategies == 0:
        best_strategy_label.config(text="No matching strategy found")
        return
    
    best_strategy_label.config(text=f"{num_matched_strategies} Matched Strategies:")
    
    # Clear previous strategy frames
    for frame in strategy_frames:
        frame.grid_forget()
    
    # Display each matched strategy in its own column
    for i, strategy in enumerate(matched_strategies):
        frame = ttk.Frame(strategies_frame)
        frame.grid(row=0, column=i, padx=5, pady=5, sticky="nsew")
        strategy_frames.append(frame)
        
        # Strategy name
        ttk.Label(frame, text=f"Strategy {i+1}:", font="TkDefaultFont 9").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        name_label = ttk.Label(frame, text=strategy['name'], font="TkDefaultFont 10 bold", foreground="black")
        name_label.grid(row=0, column=1, columnspan=2, padx=(0,5), pady=2, sticky="w")
        
        # Strategy attributes
        attributes = {
            "Difficulty": strategy['difficulty'],
            "Price Outlook": strategy['price_outlook'],
            "Volatility Outlook": strategy['volatility_outlook'],
            "Profit Potential": strategy['profit_potential'],
            "Risk": strategy['risk']
        }
        row_num = 1
        for attribute, value in attributes.items():
            ttk.Label(frame, text=attribute+":", font="TkDefaultFont 9").grid(row=row_num, column=0, padx=5, pady=(2,0), sticky="w")
            value_label = ttk.Label(frame, text=value, font="TkDefaultFont 9")
            value_label.grid(row=row_num, column=1, padx=(0,5), pady=(2,0), sticky="w")
            row_num += 1
            # Set text color based on price outlook
            if attribute == 'Price Outlook':
                if value == 'Bullish':
                    value_label.configure(foreground='green')
                elif value == 'Bearish':
                    value_label.configure(foreground='red')
                elif value == 'Neutral':
                    value_label.configure(foreground='blue')
        
        # Load and resize strategy picture
        img = Image.open(strategy['picture'])
        img = img.resize((200, 200), Image.LANCZOS)
        img = ImageTk.PhotoImage(img)
        ttk.Label(frame, image=img).grid(row=row_num, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        ttk.Label(frame, image=img).image = img
        
        # Button to create Excel file
        create_button = ttk.Button(frame, text=f"Create {strategy['name']}", command=lambda name=strategy['name']: create_excel_file(name))
        create_button.grid(row=row_num+1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")


def on_button_click():
    price_outlook = price_outlook_var.get()
    volatility_outlook = volatility_outlook_var.get()
    profit_potential = profit_potential_var.get()
    risk = risk_var.get()
    
    matched_strategies = find_matching_strategies(price_outlook, volatility_outlook, profit_potential, risk)
    display_matched_strategies(matched_strategies)

# Create main window
root = tk.Tk()
root.title("Danny's Option Strategy Selector (JPM2024)")
root.geometry("700x525")

# Create frames for matched strategies and search dropdowns
search_frame = ttk.Frame(root)
search_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

strategies_frame = ttk.Frame(root)
strategies_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

# Create dropdown menus for inputs
price_outlook_var = tk.StringVar()
volatility_outlook_var = tk.StringVar()
profit_potential_var = tk.StringVar()
risk_var = tk.StringVar()

# Dropdown choices
price_outlooks = ["Bullish", "Bearish", "Neutral"]
volatility_outlooks = ["Rising", "Falling", "Stable"]
profit_potentials = ["Limited", "Unlimited"]
risks = ["Limited", "Unlimited"]

# Create dropdowns
price_outlook_dropdown = ttk.Combobox(search_frame, textvariable=price_outlook_var, values=price_outlooks)
volatility_outlook_dropdown = ttk.Combobox(search_frame, textvariable=volatility_outlook_var, values=volatility_outlooks)
profit_potential_dropdown = ttk.Combobox(search_frame, textvariable=profit_potential_var, values=profit_potentials)
risk_dropdown = ttk.Combobox(search_frame, textvariable=risk_var, values=risks)

# Create labels for dropdowns
price_outlook_label = ttk.Label(search_frame, text="Price Outlook:")
volatility_outlook_label = ttk.Label(search_frame, text="Volatility Outlook:")
profit_potential_label = ttk.Label(search_frame, text="Profit Potential:")
risk_label = ttk.Label(search_frame, text="Risk:")

# Create button to find best strategy
find_strategy_button = ttk.Button(search_frame, text="Find Best Strategy", command=on_button_click)

# Create label to display best strategy
best_strategy_label = ttk.Label(strategies_frame, text="")

# Layout
search_frame.columnconfigure(1, weight=1)  # Expand column to fill space
price_outlook_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
price_outlook_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
volatility_outlook_label.grid(row=0, column=2, padx=5, pady=5, sticky="e")
volatility_outlook_dropdown.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
profit_potential_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
profit_potential_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
risk_label.grid(row=1, column=2, padx=5, pady=5, sticky="e")
risk_dropdown.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
find_strategy_button.grid(row=2, column=3, padx=5, pady=5, sticky="e")
best_strategy_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

# Initialize list to hold strategy frames
strategy_frames = []

# Run the application
root.mainloop()
