import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import configparser
import json
import os
from pathlib import Path
from typing import Dict, List, Optional

class ConfigUI:
    def __init__(self, root: Optional[tk.Tk] = None):
        """Initialize the configuration UI."""
        self.root = root or tk.Tk()
        self.root.title("EtereBridge Configuration")
        self.root.geometry("800x600")
        
        # Store the current configuration
        self.config = configparser.ConfigParser()
        self.config.optionxform = str  # Preserve case sensitivity
        
        # Create main container with padding
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create tabs
        self.paths_tab = self._create_paths_tab()
        self.markets_tab = self._create_markets_tab()
        self.sales_tab = self._create_sales_tab()
        self.columns_tab = self._create_columns_tab()
        
        # Add tabs to notebook
        self.notebook.add(self.paths_tab, text="Paths")
        self.notebook.add(self.markets_tab, text="Markets")
        self.notebook.add(self.sales_tab, text="Sales")
        self.notebook.add(self.columns_tab, text="Columns")
        
        # Add save/load buttons at bottom
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=1, column=0, pady=10)
        
        ttk.Button(self.button_frame, text="Load Config", 
                  command=self.load_config).grid(row=0, column=0, padx=5)
        ttk.Button(self.button_frame, text="Save Config", 
                  command=self.save_config).grid(row=0, column=1, padx=5)
        ttk.Button(self.button_frame, text="Reset to Defaults", 
                  command=self.reset_to_defaults).grid(row=0, column=2, padx=5)

    def _create_paths_tab(self) -> ttk.Frame:
        """Create the Paths configuration tab."""
        frame = ttk.Frame(self.notebook, padding="10")
        
        # Path entries with browse buttons
        paths = [
            ("Template Path", "template_path"),
            ("Input Directory", "input_dir"),
            ("Output Directory", "output_dir")
        ]
        
        self.path_vars = {}
        for i, (label, key) in enumerate(paths):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=5)
            
            # Create variable and entry
            var = tk.StringVar()
            self.path_vars[key] = var
            entry = ttk.Entry(frame, textvariable=var, width=50)
            entry.grid(row=i, column=1, padx=5)
            
            # Add browse button
            ttk.Button(frame, text="Browse", 
                      command=lambda k=key: self._browse_path(k)).grid(row=i, column=2)
        
        return frame

    def _create_markets_tab(self) -> ttk.Frame:
        """Create the Markets configuration tab."""
        frame = ttk.Frame(self.notebook, padding="10")
        
        # Market mappings table
        self.market_tree = ttk.Treeview(frame, columns=("Original", "Replacement"), 
                                      show="headings")
        self.market_tree.heading("Original", text="Original Market Name")
        self.market_tree.heading("Replacement", text="Replacement Name")
        self.market_tree.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for table
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.market_tree.yview)
        scrollbar.grid(row=0, column=3, sticky=(tk.N, tk.S))
        self.market_tree.configure(yscrollcommand=scrollbar.set)
        
        # Add/Edit/Delete buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=4, pady=10)
        
        ttk.Button(btn_frame, text="Add Market", 
                  command=self._add_market).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Edit Market", 
                  command=self._edit_market).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Delete Market", 
                  command=self._delete_market).grid(row=0, column=2, padx=5)
        
        return frame

    def _create_sales_tab(self) -> ttk.Frame:
        """Create the Sales configuration tab."""
        frame = ttk.Frame(self.notebook, padding="10")
        
        # Sales people list
        self.sales_tree = ttk.Treeview(frame, columns=("Name",), show="headings")
        self.sales_tree.heading("Name", text="Sales Person Name")
        self.sales_tree.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.sales_tree.yview)
        scrollbar.grid(row=0, column=2, sticky=(tk.N, tk.S))
        self.sales_tree.configure(yscrollcommand=scrollbar.set)
        
        # Add/Delete buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=10)
        
        ttk.Button(btn_frame, text="Add Sales Person", 
                  command=self._add_sales_person).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Delete Sales Person", 
                  command=self._delete_sales_person).grid(row=0, column=1, padx=5)
        
        return frame

    def _create_columns_tab(self) -> ttk.Frame:
        """Create the Columns configuration tab."""
        frame = ttk.Frame(self.notebook, padding="10")
        
        # Column order list
        self.columns_tree = ttk.Treeview(frame, columns=("Column",), show="headings")
        self.columns_tree.heading("Column", text="Column Name")
        self.columns_tree.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.columns_tree.yview)
        scrollbar.grid(row=0, column=2, sticky=(tk.N, tk.S))
        self.columns_tree.configure(yscrollcommand=scrollbar.set)
        
        # Move up/down buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=10)
        
        ttk.Button(btn_frame, text="Move Up", 
                  command=self._move_column_up).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Move Down", 
                  command=self._move_column_down).grid(row=0, column=1, padx=5)
        
        return frame

    def _browse_path(self, key: str):
        """Open file/directory browser dialog."""
        if key == "template_path":
            path = filedialog.askopenfilename(
                title="Select Template File",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
        else:
            path = filedialog.askdirectory(title=f"Select {key} Directory")
            
        if path:
            self.path_vars[key].set(path)

    def _add_market(self):
        """Add a new market mapping."""
        dialog = MarketDialog(self.root)
        if dialog.result:
            original, replacement = dialog.result
            self.market_tree.insert("", tk.END, values=(original, replacement))

    def _edit_market(self):
        """Edit selected market mapping."""
        selected = self.market_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a market to edit.")
            return
            
        current = self.market_tree.item(selected[0])["values"]
        dialog = MarketDialog(self.root, current)
        if dialog.result:
            self.market_tree.item(selected[0], values=dialog.result)

    def _delete_market(self):
        """Delete selected market mapping."""
        selected = self.market_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a market to delete.")
            return
            
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this market mapping?"):
            self.market_tree.delete(selected[0])

    def load_config(self):
        """Load configuration from file."""
        try:
            self.config.read("config.ini")
            
            # Load paths
            for key in self.path_vars:
                if key in self.config["Paths"]:
                    self.path_vars[key].set(self.config["Paths"][key])
            
            # Load markets
            self.market_tree.delete(*self.market_tree.get_children())
            for orig, repl in self.config["Markets"].items():
                self.market_tree.insert("", tk.END, values=(orig, repl))
            
            # Load sales people
            self.sales_tree.delete(*self.sales_tree.get_children())
            sales_people = self.config["Sales"]["sales_people"].split(",")
            for person in sales_people:
                self.sales_tree.insert("", tk.END, values=(person.strip(),))
            
            # Load columns
            self.columns_tree.delete(*self.columns_tree.get_children())
            columns = self.config["Columns"]["final_columns"].split(",")
            for col in columns:
                self.columns_tree.insert("", tk.END, values=(col.strip(),))
            
            messagebox.showinfo("Success", "Configuration loaded successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")

    def save_config(self):
        """Save configuration to file."""
        try:
            # Save paths
            self.config["Paths"] = {
                key: var.get() for key, var in self.path_vars.items()
            }
            
            # Save markets
            self.config["Markets"] = {
                item["values"][0]: item["values"][1]
                for item in self.market_tree.get_children()
            }
            
            # Save sales people
            sales_people = [
                self.sales_tree.item(item)["values"][0]
                for item in self.sales_tree.get_children()
            ]
            self.config["Sales"] = {
                "sales_people": ",".join(sales_people)
            }
            
            # Save columns
            columns = [
                self.columns_tree.item(item)["values"][0]
                for item in self.columns_tree.get_children()
            ]
            self.config["Columns"] = {
                "final_columns": ",".join(columns)
            }
            
            with open("config.ini", "w") as f:
                self.config.write(f)
            
            messagebox.showinfo("Success", "Configuration saved successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")

    def run(self):
        """Start the configuration UI."""
        self.root.mainloop()


class MarketDialog:
    """Dialog for adding/editing market mappings."""
    def __init__(self, parent, current=None):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Market Mapping")
        self.result = None
        
        # Create and layout widgets
        ttk.Label(self.dialog, text="Original Market Name:").grid(row=0, column=0, pady=5)
        ttk.Label(self.dialog, text="Replacement Name:").grid(row=1, column=0, pady=5)
        
        self.orig_var = tk.StringVar(value=current[0] if current else "")
        self.repl_var = tk.StringVar(value=current[1] if current else "")
        
        ttk.Entry(self.dialog, textvariable=self.orig_var).grid(row=0, column=1, padx=5)
        ttk.Entry(self.dialog, textvariable=self.repl_var).grid(row=1, column=1, padx=5)
        
        ttk.Button(self.dialog, text="OK", command=self._ok).grid(row=2, column=0, pady=10)
        ttk.Button(self.dialog, text="Cancel", command=self.dialog.destroy).grid(row=2, column=1)
        
        self.dialog.wait_window()

    def _ok(self):
        """Handle OK button click."""
        orig = self.orig_var.get().strip()
        repl = self.repl_var.get().strip()
        
        if orig and repl:
            self.result = (orig, repl)
            self.dialog.destroy()
        else:
            messagebox.showwarning("Invalid Input", "Both fields are required.")


if __name__ == "__main__":
    ui = ConfigUI()
    ui.run()