import sys
import textwrap
import numpy as np
import pandas as pd
import matplotlib.ticker
from PyQt5 import QtCore
import matplotlib.pyplot as plt
import matplotlib.colors as colors
from PyQt5.QtCore import Qt, QDate
plt.rcParams.update({'font.size': 22})
from openpyxl import Workbook, load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QDateEdit, QLineEdit, QPushButton, QMainWindow, QTextBrowser

class InputWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.data = pd.DataFrame(columns=["Category", "Date", "Price", "Description"])
        self.init_ui()
        
    def truncate_colormap(self, cmap, min_val=0.0, max_val=1.0, n=100):
        # Truncate the color map according to the min_val and max_val from the original color map.
        new_cmap = colors.LinearSegmentedColormap.from_list('trunc({n},{a:.2f},{b:.2f})'.format(n=cmap.name, a=min_val, b=max_val),cmap(np.linspace(min_val, max_val, n)))
        return new_cmap

    def init_ui(self):
        # Set window size (X, Y, WIDTH, HEIGHT)
        self.setGeometry(1400, 700, 900, 1100)
        layout = QVBoxLayout()
        line1_layout = QHBoxLayout()

        category_label = QLabel("Category:")
        self.category_combo = QComboBox()
        self.category_combo.addItems(["Shirdi Trip", "Water or Milk or Curd", "Fruits & Vegitables", "Food & Snacks", "Groceries", "Households & Personal Care", "Others", "Tickets", "Petrol or Rapido or Cab", "Rent", "Investments", "Home Appliances", "Recharges & Bills", "Shopping", "Medicine & Medical Bills"])
        layout.addWidget(category_label)
        layout.addWidget(self.category_combo)

        date_label = QLabel("Date:")
        self.date_input = QDateEdit()
        self.date_input.setDate(QtCore.QDate.currentDate())
        layout.addWidget(date_label)
        layout.addWidget(self.date_input)

        price_label = QLabel("Price:")
        self.price_input = QLineEdit()
        layout.addWidget(price_label)
        layout.addWidget(self.price_input)
        
        description_label = QLabel("Description:")
        self.description_combo = QComboBox()
        self.description_combo.addItems([" ","Tickets", "Local_transport", "Food", "Accommodation", "Lunch",
                                         "Pizza or Burger", "Lunch & Tea", "Milk", "Curd", "Milk & Curd", "Watercan", "Petrol", "Snacks in Grace", "Parking Charges", "Others"])
        self.description_input = QLineEdit()
        layout.addWidget(description_label)
        layout.addWidget(self.description_combo)
        layout.addWidget(self.description_input)

        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_data)
        layout.addWidget(save_button)

        cost_analysis_button = QPushButton("Expences Ratios")
        cost_analysis_button.clicked.connect(self.show_cost_analysis)
        self.cost_analysis_combo = QComboBox()
        self.cost_analysis_combo.addItems(["June", "July", "August", "September", "October", "November", "December"])
        # Set the current month as the default selection
        current_month_index = QDate.currentDate().month() - 6
        self.cost_analysis_combo.setCurrentIndex(current_month_index)
        line1_layout.addWidget(cost_analysis_button)
        line1_layout.addWidget(self.cost_analysis_combo)
        layout.addLayout(line1_layout)
        
        cost_analysis_button2 = QPushButton("Expences in Bars")
        cost_analysis_button2.clicked.connect(self.show_expences_in_bars)
        layout.addWidget(cost_analysis_button2)
        
        self.button_last_10 = QPushButton("Last 5 Entries")
        self.text_browser = QTextBrowser()
        layout.addWidget(self.button_last_10)
        layout.addWidget(self.text_browser)
        self.button_last_10.clicked.connect(self.show_last_10)

        self.setLayout(layout)
        self.setWindowTitle("Price Analysis")
        self.show()

    def save_data(self):
        category = self.category_combo.currentText()
        date = self.date_input.date().toString("yyyy-MM-dd")
        price = int(self.price_input.text())
        Details = self.description_input.text() if self.description_input.text() != "" else self.description_combo.currentText()

        # Append data to DataFrame
        new_data = pd.DataFrame({"Category": [category], "Date": [date], "Price": [price], "Description":[Details]})
        self.data = pd.concat([self.data, new_data], ignore_index=True)

        # Write data to Excel
        workbook = load_workbook("Price Analysis.xlsx")
        sheet = workbook.active
        sheet.append([category, Details, date, price])
        workbook.save("Price Analysis.xlsx")
        
        # Clear the QLineEdit
        self.price_input.clear()
        self.description_input.clear()
        # Reset the QComboBox
        self.category_combo.setCurrentIndex(0)
        self.description_combo.setCurrentIndex(0)
       
    def show_cost_analysis(self):
        # Calculate total price by category
        data_to_plot = pd.read_excel("Price Analysis.xlsx", sheet_name = self.cost_analysis_combo.currentText())
        total_price = data_to_plot.groupby("Category")["Price"].sum()
        total_price = total_price.sort_values()
        condition = data_to_plot.iloc[:, 0] == 'Shirdi Trip'
        if condition.any():
            selected_rows = data_to_plot[condition]
            total_price2 = selected_rows.groupby("Description")["Price"].sum()
            fig = plt.figure(figsize=(38.5, 21.5))  # Adjust the values (width, height) as per your preference
        else:
            # Create a figure with a larger size
            fig = plt.figure(figsize=(80, 40))  # Adjust the values (width, height) as per your preference

        # Create a pie chart
        if condition.any():
            plt.subplot(1, 2, 1)  # Create a subplot for the main pie chart
        plt.pie(total_price, labels=total_price.index, autopct='%1.1f%%')
        plt.title("Total Price by Category")
        legend_values = [f'{c} = {v}' for c, v in zip(total_price.index, total_price.values)]
        plt.legend(legend_values)
        # Add total as text annotation
        plt.annotate(f'Total: {sum(total_price.values)}', (1.2, 0.2), xycoords='axes fraction', va='center', ha='center', fontsize=36)

        if condition.any():
            # Plot the sub pie chart for the "Trip" category
            plt.subplot(1, 2, 2)  # Create a subplot for the "Trip" pie chart
            plt.pie(total_price2, labels=total_price2.index, autopct='%1.1f%%')
            plt.title("Shirdi Trip Expenses")

            # Adjust the layout to avoid overlapping
            plt.tight_layout()
        # Show the plot
        plt.show()
        
    def show_expences_in_bars(self):
        # Calculate total price by category
        data_to_plot = pd.read_excel("Price Analysis.xlsx", sheet_name = self.cost_analysis_combo.currentText())
        total_price = data_to_plot.groupby("Category")["Price"].sum()
        
        # Define colors based on threshold limits
        # colors = ['red' if value >= 5000 else 'goldenrod' if value >= 2000 else "orange" if value >= 1500 else 'greenyellow' if value >= 1000 else 'green' for value in total_price.values]
        
        # Create a plot
        fig, ax = plt.subplots(figsize=(37, 21))
        # bar_plot = ax.bar(total_price.index, total_price.values, width = 0.3, color = colors)
        bar_plot = ax.bar(total_price.index, total_price.values, width = 0.3)
        x_min, x_max = ax.get_xlim()
        # Get the minimum and maximum values of the y-axis
        y_min, y_max = ax.get_ylim()
        y_min_top = min(total_price.values)
        
        # Create a gradient array
        grad = np.atleast_2d(np.linspace(0, 1, 256)).T
        
        for bar in bar_plot:
            bar.set_zorder(1)  # put the bars in front
            bar.set_facecolor("none")  # make the bars transparent
            x, _ = bar.get_xy()  # get the corners
            w, h = bar.get_width(), bar.get_height()  # get the width and height

            # Calculate the height of the bar as a percentage of the y-axis range
            bar_height_percentage = (h - y_min) / (y_max - y_min)

            # Truncate the jet colormap based on the bar height percentage
            c_map = self.truncate_colormap(plt.cm.jet, min_val=0, max_val=bar_height_percentage)

            # Overlay the truncated gradient on the bar using imshow
            ax.imshow(grad, extent=[x, x+w, h, y_min], aspect="auto", zorder=0, cmap=c_map)

        # Set the axis limits to include the bars and the gradient overlay
        # print ("1. ", x_min, x_max, y_min, y_max)
        ax.axis([x_min, x_max, y_min, y_max])
        
        plt.title("Total Price by Category")
        plt.ylabel("Expence in rupees")
        # Wrap the labels into two lines
        wrapped_labels = [textwrap.fill(label, 10) for label in total_price.index]
        # Set the wrapped labels as x-tick labels
        plt.xticks(range(len(total_price.index)), wrapped_labels)
        
        # Add value annotations to each bar
        for i, v in enumerate(total_price.values):
            plt.text(i, v, str(v), ha='center', va='bottom')
        # Add total as text annotation
        plt.annotate(f'Total: {sum(total_price.values)}', (0.1, 0.8), xycoords='axes fraction', va='center', ha='center', fontsize=36)
        
        # Define the function to toggle the y-axis scaling
        def toggle_y_axis_scaling(event):
            if event.key == 't':
                if ax.get_yscale() == 'log':
                    ax.set_yscale('linear')
                    ax.axis([x_min, x_max, y_min, y_max])
                    ax.set_ylabel('Expence in rupees (linear scale)')
                else:
                    ax.set_yscale('log')
                    # Set the axis limits to include the bars and the gradient overlay
                    # print (x_min, x_max, y_min, y_max)
                    ax.axis([x_min, x_max, 1 if y_min_top < 11 else y_min_top-10 , y_max+10])
                    ax.set_yticks([100, 200, 300, 400, 500, 700, 1000, 2000, 3000, 4000, 5000, 7000, 10000])
                    ax.get_yaxis().set_major_formatter(matplotlib.ticker.ScalarFormatter())
                    ax.set_ylabel('Expence in rupees (log scale)')
                plt.draw()

        # Connect the key press event to the toggle function
        fig.canvas.mpl_connect('key_press_event', toggle_y_axis_scaling)
        
        # Show the plot
        plt.show()
        
    def show_last_10(self):
        data_to_show = pd.read_excel("Price Analysis.xlsx", sheet_name = self.cost_analysis_combo.currentText())
        last_10_entries = data_to_show.tail(5)
        last_10_entries = last_10_entries.iloc[:,:4]
        self.text_browser.clear()
        self.text_browser.setPlainText(last_10_entries.to_string(index=False))
                

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = InputWindow()
    sys.exit(app.exec_())
