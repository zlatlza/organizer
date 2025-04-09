import os
import json
import re
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
import PyPDF2
import subprocess
import platform
import dateutil.parser

class CategoryEditor(tk.Toplevel):
    def __init__(self, parent, categories, callback):
        super().__init__(parent)
        self.title("Category Editor")
        self.geometry("700x550")  # Slightly larger for better usability
        self.resizable(True, True)
        self.parent = parent
        self.categories = categories.copy()
        self.callback = callback
        
        # Create style for visual enhancements
        self.style = ttk.Style()
        self.style.configure("TButton", padding=3)
        self.style.configure("Add.TButton", foreground="green")
        self.style.configure("Remove.TButton", foreground="red")
        
        # Set initial variables
        self.current_category = None
        self.previous_selection = None  # Track previous selection to detect changes
        self.keyword_entries = []
        self.has_unsaved_changes = False
        self.currently_editing = False  # Flag to track if we're in edit mode
        self.ignore_selection_change = False  # Flag to ignore temporary selection changes
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main container with padding
        self.main_container = ttk.Frame(self, padding="10")
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid for main container
        self.main_container.columnconfigure(0, weight=1)
        self.main_container.columnconfigure(1, weight=1)
        self.main_container.rowconfigure(0, weight=1)
        
        # Left panel - Category list and controls
        self.left_panel = ttk.LabelFrame(self.main_container, text="Categories", padding="5")
        self.left_panel.grid(row=0, column=0, padx=(0, 5), pady=5, sticky="nsew")
        
        # Category search
        self.search_frame = ttk.Frame(self.left_panel)
        self.search_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_categories)
        
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.search_clear = ttk.Button(self.search_frame, text="×", width=3, 
                                      command=lambda: self.search_var.set(""))
        self.search_clear.pack(side=tk.LEFT, padx=(2, 0))
        
        # Category listbox with scrollbar in a frame
        self.list_frame = ttk.Frame(self.left_panel)
        self.list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.category_listbox = tk.Listbox(self.list_frame, activestyle="dotbox", 
                                          selectbackground="#007bff", selectforeground="white",
                                          exportselection=False)  # Prevent selection from being cleared
        self.category_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.category_listbox.bind('<<ListboxSelect>>', self.on_category_select)
        self.category_listbox.bind('<Double-Button-1>', self.edit_category)
        
        self.scrollbar = ttk.Scrollbar(self.list_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Connect scrollbar to listbox
        self.category_listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.category_listbox.yview)
        
        # Buttons for category management
        self.button_frame = ttk.Frame(self.left_panel)
        self.button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.add_button = ttk.Button(self.button_frame, text="Add New", style="Add.TButton", 
                                    command=self.add_category)
        self.add_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        
        self.edit_button = ttk.Button(self.button_frame, text="Rename", 
                                     command=self.edit_category)
        self.edit_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.remove_button = ttk.Button(self.button_frame, text="Remove", style="Remove.TButton",
                                       command=self.remove_category)
        self.remove_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))
        
        # Right panel - Category details
        self.right_panel = ttk.LabelFrame(self.main_container, text="Category Details", padding="10")
        self.right_panel.grid(row=0, column=1, padx=(5, 0), pady=5, sticky="nsew")
        
        # Configure grid for details panel
        self.right_panel.columnconfigure(0, weight=0)
        self.right_panel.columnconfigure(1, weight=1)
        self.right_panel.rowconfigure(3, weight=1)  # Keywords section expands
        
        # Category name display (not editable here - use rename button)
        ttk.Label(self.right_panel, text="Name:").grid(row=0, column=0, padx=5, pady=8, sticky=tk.W)
        
        self.name_var = tk.StringVar()
        self.name_display = ttk.Label(self.right_panel, textvariable=self.name_var, font=("", 10, "bold"))
        self.name_display.grid(row=0, column=1, padx=5, pady=8, sticky=tk.W+tk.E)
        
        # Folder name
        ttk.Label(self.right_panel, text="Folder:").grid(row=1, column=0, padx=5, pady=8, sticky=tk.W)
        
        self.folder_var = tk.StringVar()
        self.folder_frame = ttk.Frame(self.right_panel)
        self.folder_frame.grid(row=1, column=1, padx=5, pady=8, sticky=tk.W+tk.E)
        
        self.folder_entry = ttk.Entry(self.folder_frame, textvariable=self.folder_var)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.folder_var.trace_add("write", self.on_field_change)
        
        # Set focus handler for all entry widgets
        self.folder_entry.bind("<FocusIn>", self.on_field_focus)
        
        # Add button to auto-capitalize folder name
        self.capitalize_button = ttk.Button(self.folder_frame, text="Auto", width=5,
                                          command=self.auto_capitalize_folder)
        self.capitalize_button.pack(side=tk.LEFT, padx=(3, 0))
        
        # Abbreviation
        ttk.Label(self.right_panel, text="Abbreviation:").grid(row=2, column=0, padx=5, pady=8, sticky=tk.W)
        
        self.abbr_var = tk.StringVar()
        self.abbr_frame = ttk.Frame(self.right_panel)
        self.abbr_frame.grid(row=2, column=1, padx=5, pady=8, sticky=tk.W+tk.E)
        
        self.abbr_entry = ttk.Entry(self.abbr_frame, textvariable=self.abbr_var)
        self.abbr_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.abbr_var.trace_add("write", self.on_field_change)
        self.abbr_entry.bind("<FocusIn>", self.on_field_focus)
        
        # Add button to auto-generate abbreviation
        self.auto_abbr_button = ttk.Button(self.abbr_frame, text="Auto", width=5,
                                         command=self.auto_generate_abbreviation)
        self.auto_abbr_button.pack(side=tk.LEFT, padx=(3, 0))
        
        # Keywords section
        ttk.Label(self.right_panel, text="Keywords:").grid(row=3, column=0, padx=5, pady=(8, 0), sticky=tk.NW)
        
        # Keywords container
        self.keywords_container = ttk.Frame(self.right_panel)
        self.keywords_container.grid(row=3, column=1, padx=5, pady=(8, 0), sticky="nsew")
        self.keywords_container.columnconfigure(0, weight=1)
        
        # Scrollable frame for keywords
        self.canvas = tk.Canvas(self.keywords_container, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.keywords_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", tags="self.scrollable_frame")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Allow the scrollable area to expand
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Enable mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Bind click event to canvas to maintain selection
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        
        # Add keyword button
        self.add_keyword_frame = ttk.Frame(self.right_panel)
        self.add_keyword_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=8, sticky=tk.E+tk.W)
        
        self.new_keyword_var = tk.StringVar()
        self.new_keyword_entry = ttk.Entry(self.add_keyword_frame, textvariable=self.new_keyword_var)
        self.new_keyword_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))
        self.new_keyword_entry.bind("<Return>", lambda e: self.add_keyword())
        self.new_keyword_entry.bind("<FocusIn>", self.on_field_focus)
        
        self.add_keyword_button = ttk.Button(self.add_keyword_frame, text="Add Keyword", 
                                           command=self.add_keyword)
        self.add_keyword_button.pack(side=tk.RIGHT)
        
        # Save details button (for each category)
        self.save_details_button = ttk.Button(self.right_panel, text="Apply Changes", 
                                            command=self.save_details)
        self.save_details_button.grid(row=5, column=0, columnspan=2, pady=(10, 0), sticky=tk.E)
        
        # Bottom buttons (Save/Cancel)
        self.bottom_frame = ttk.Frame(self)
        self.bottom_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Status message
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self.bottom_frame, textvariable=self.status_var, 
                                     foreground="green")
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # Buttons
        self.cancel_button = ttk.Button(self.bottom_frame, text="Cancel", 
                                       command=self.confirm_cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=5)
        
        self.save_button = ttk.Button(self.bottom_frame, text="Save All Changes", 
                                     command=self.save_changes)
        self.save_button.pack(side=tk.RIGHT, padx=5)
        
        # Populate the list and set initial state
        self.populate_categories()
        
        # Set initial selection if categories exist
        if self.category_listbox.size() > 0:
            self.category_listbox.selection_set(0)  # Select first item
            self.category_listbox.event_generate('<<ListboxSelect>>')  # Trigger selection event
        else:
            self.update_ui_state()  # Update UI state for empty list
        
        # Make dialog modal
        self.transient(self.parent)
        self.grab_set()
        
        # Focus on search field initially
        self.search_entry.focus_set()
        
        # Handle window close event
        self.protocol("WM_DELETE_WINDOW", self.confirm_cancel)
        
        # Bind events to maintain selection
        self.bind("<Button-1>", self.ensure_selection_maintained)
    
    def ensure_selection_maintained(self, event=None):
        """Ensure the current category selection is maintained"""
        if self.current_category and not self.category_listbox.curselection():
            # Find the item in the listbox
            for i in range(self.category_listbox.size()):
                if self.category_listbox.get(i) == self.current_category:
                    self.category_listbox.selection_set(i)
                    self.ignore_selection_change = True
                    break
        # Allow the event to propagate
        return
    
    def on_field_focus(self, event=None):
        """Handle focus on input fields to maintain selection"""
        self.ensure_selection_maintained()
    
    def on_canvas_click(self, event=None):
        """Handle clicks on the canvas to maintain selection"""
        self.ensure_selection_maintained()
    
    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling in the keywords area"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def on_field_change(self, *args):
        """Track changes to form fields"""
        self.has_unsaved_changes = True
    
    def filter_categories(self, *args):
        """Filter the categories list based on the search text"""
        search_text = self.search_var.get().lower()
        
        self.category_listbox.delete(0, tk.END)
        
        for category in sorted(self.categories.keys()):
            if search_text in category.lower():
                self.category_listbox.insert(tk.END, category)
        
        # Handle selection after filtering
        if self.category_listbox.size() > 0:
            # Try to select previously selected category if it's still in the filtered list
            found = False
            if self.current_category:
                for i in range(self.category_listbox.size()):
                    if self.category_listbox.get(i) == self.current_category:
                        self.category_listbox.selection_set(i)
                        self.category_listbox.see(i)
                        found = True
                        break
                        
            # If previous selection not found, select first item
            if not found:
                self.category_listbox.selection_set(0)
                self.current_category = self.category_listbox.get(0)
                
            # Update UI with selected category
            self.on_category_select(None)
        else:
            # No categories match filter
            self.current_category = None
            self.update_ui_state()
    
    def on_category_select(self, event):
        """Handle category selection"""
        # If there's no selection or we should ignore the change, do nothing
        if not self.category_listbox.curselection() or self.ignore_selection_change:
            if self.ignore_selection_change:
                self.ignore_selection_change = False
            self.update_ui_state()
            return
        
        selected_index = self.category_listbox.curselection()[0]
        selected_category = self.category_listbox.get(selected_index)
        
        # Check if this is the same category already selected (do nothing to avoid interrupting edits)
        if selected_category == self.current_category:
            return
            
        # Check if there are unsaved changes before switching
        if self.has_unsaved_changes and self.current_category:
            if messagebox.askyesno("Unsaved Changes", 
                                  f"You have unsaved changes to '{self.current_category}'. Save them?"):
                self.save_details()
        
        # Clear the status message
        self.status_var.set("")
        
        # Update current category
        self.current_category = selected_category
        
        # Update details fields
        self.name_var.set(selected_category)
        self.folder_var.set(self.categories[selected_category].get("folder", selected_category.capitalize()))
        self.abbr_var.set(self.categories[selected_category].get("abbreviation", selected_category.upper()[:4]))
        
        # Update keywords
        self.refresh_keywords()
        
        # Reset unsaved changes flag
        self.has_unsaved_changes = False
        
        # Update UI state
        self.update_ui_state()
    
    def refresh_keywords(self):
        """Refresh the keywords display"""
        # Save scroll position
        current_scroll = self.canvas.yview()
        
        # Clear existing keywords UI
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.keyword_entries = []
        
        if not self.current_category:
            return
            
        # Get keywords
        keywords = self.categories[self.current_category].get("keywords", [])
        
        # Add keyword entries
        for i, keyword in enumerate(keywords):
            keyword_frame = ttk.Frame(self.scrollable_frame)
            keyword_frame.pack(fill=tk.X, pady=2)
            
            # Keyword index label
            index_label = ttk.Label(keyword_frame, text=f"{i+1}.", width=3)
            index_label.pack(side=tk.LEFT, padx=(0, 5))
            
            # Keyword entry
            keyword_var = tk.StringVar(value=keyword)
            keyword_entry = ttk.Entry(keyword_frame, textvariable=keyword_var)
            keyword_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
            keyword_var.trace_add("write", self.on_field_change)
            keyword_entry.bind("<FocusIn>", self.on_field_focus)
            
            # Delete button
            delete_button = ttk.Button(keyword_frame, text="×", width=3,
                                     command=lambda idx=i: self.remove_keyword(idx))
            delete_button.pack(side=tk.RIGHT)
            
            self.keyword_entries.append(keyword_var)
        
        # Reset canvas scroll region
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Restore scroll position if possible
        try:
            self.canvas.yview_moveto(current_scroll[0])
        except:
            pass
    
    def add_keyword(self):
        """Add a new keyword to the current category"""
        if not self.current_category:
            return
            
        keyword = self.new_keyword_var.get().strip()
        if not keyword:
            return
            
        # Add keyword to the list
        if "keywords" not in self.categories[self.current_category]:
            self.categories[self.current_category]["keywords"] = []
            
        current_keywords = self.categories[self.current_category]["keywords"]
        
        # Check if keyword already exists
        if keyword in current_keywords:
            # Just clear the field
            self.new_keyword_var.set("")
            return
            
        # Add the keyword to the current category
        current_keywords.append(keyword)
        
        # Mark changes
        self.has_unsaved_changes = True
        
        # Update UI without changing selection
        self.refresh_keywords()
        
        # Clear the new keyword field
        self.new_keyword_var.set("")
        
        # Focus back to the new keyword field
        self.new_keyword_entry.focus_set()
        
        # Show success message
        self.status_var.set(f"Keyword '{keyword}' added")
        
        # Make sure selection is maintained
        self.ensure_selection_maintained()
    
    def remove_keyword(self, index):
        """Remove a keyword at the specified index"""
        if not self.current_category:
            return
            
        if "keywords" not in self.categories[self.current_category]:
            return
            
        current_keywords = self.categories[self.current_category]["keywords"]
        if index < 0 or index >= len(current_keywords):
            return
            
        # Remove the keyword
        removed_keyword = current_keywords.pop(index)
        
        # Mark changes
        self.has_unsaved_changes = True
        
        # Update UI without changing selection
        self.refresh_keywords()
        
        # Show success message
        self.status_var.set(f"Keyword '{removed_keyword}' removed")
        
        # Make sure selection is maintained
        self.ensure_selection_maintained()
    
    def save_details(self):
        """Save current category details"""
        if not self.current_category:
            return False
            
        # Get values
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showinfo("Info", "Folder name cannot be empty")
            self.folder_entry.focus_set()
            return False
            
        abbreviation = self.abbr_var.get().strip()
        if not abbreviation:
            messagebox.showinfo("Info", "Abbreviation cannot be empty")
            self.abbr_entry.focus_set()
            return False
            
        # Get keywords from entries
        keywords = [var.get().strip() for var in self.keyword_entries]
        keywords = [k for k in keywords if k]  # Remove empty strings
        
        # Update category details
        self.categories[self.current_category]["folder"] = folder
        self.categories[self.current_category]["abbreviation"] = abbreviation
        self.categories[self.current_category]["keywords"] = keywords
        
        # Reset unsaved changes flag
        self.has_unsaved_changes = False
        
        # Confirm update
        self.status_var.set(f"Details for '{self.current_category}' saved")
        
        return True
    
    def confirm_cancel(self):
        """Ask for confirmation before closing if there are unsaved changes"""
        if self.has_unsaved_changes:
            if messagebox.askyesno("Unsaved Changes", 
                                  "You have unsaved changes. Discard changes and close?"):
                self.destroy()
        else:
            self.destroy()
    
    def save_changes(self):
        """Save all changes and close dialog"""
        if self.current_category and self.has_unsaved_changes:
            # Save current category details before closing
            if not self.save_details():
                return
        
        # Call the callback with the updated categories
        self.callback(self.categories)
        
        # Close the dialog
        self.destroy()

    def populate_categories(self):
        """Populate the categories listbox"""
        self.category_listbox.delete(0, tk.END)
        
        for category in sorted(self.categories.keys()):
            self.category_listbox.insert(tk.END, category)
    
    def update_ui_state(self):
        """Update the state of UI elements based on selection"""
        has_categories = self.category_listbox.size() > 0
        has_selection = bool(self.category_listbox.curselection())
        
        # Category management buttons
        self.edit_button.config(state="normal" if has_selection else "disabled")
        self.remove_button.config(state="normal" if has_selection else "disabled")
        
        # Details fields
        details_state = "normal" if has_selection else "disabled"
        
        # Only disable entry fields if no categories at all - otherwise keep them enabled
        # for better UX even when switching between categories
        self.folder_entry.config(state=details_state)
        self.abbr_entry.config(state=details_state)
        self.new_keyword_entry.config(state=details_state)
        self.add_keyword_button.config(state=details_state)
        self.capitalize_button.config(state=details_state)
        self.auto_abbr_button.config(state=details_state)
        self.save_details_button.config(state=details_state)
        
        # Set name field
        if has_selection:
            self.name_var.set(self.current_category)
        else:
            self.name_var.set("No category selected")
    
    def add_category(self):
        """Add a new category"""
        # Check if there are unsaved changes first
        if self.has_unsaved_changes and self.current_category:
            if messagebox.askyesno("Unsaved Changes", 
                                  f"You have unsaved changes to '{self.current_category}'. Save them?"):
                self.save_details()
        
        new_name = simpledialog.askstring("New Category", "Enter category name:", parent=self)
        if not new_name or not new_name.strip():
            return
            
        new_name = new_name.strip()
            
        # Check if category already exists
        if new_name in self.categories:
            messagebox.showerror("Error", f"Category '{new_name}' already exists")
            return
            
        # Add new category with default values
        self.categories[new_name] = {
            "folder": new_name.capitalize(),
            "abbreviation": new_name.upper()[:4],
            "keywords": []
        }
        
        # Update the list
        self.populate_categories()
        
        # Find and select the new category
        for i in range(self.category_listbox.size()):
            if self.category_listbox.get(i) == new_name:
                self.category_listbox.selection_clear(0, tk.END)
                self.category_listbox.selection_set(i)
                self.category_listbox.see(i)
                self.current_category = new_name
                break
                
        # Update UI
        self.name_var.set(new_name)
        self.folder_var.set(self.categories[new_name].get("folder", new_name.capitalize()))
        self.abbr_var.set(self.categories[new_name].get("abbreviation", new_name.upper()[:4]))
        self.refresh_keywords()
        self.update_ui_state()
        
        # Reset unsaved changes flag for the new category
        self.has_unsaved_changes = False
                
        # Set focus to folder field to continue editing
        self.folder_entry.focus_set()
        
        # Show success message
        self.status_var.set(f"Category '{new_name}' created")
    
    def edit_category(self, event=None):
        """Edit the selected category name"""
        # Check if there are unsaved changes first
        if self.has_unsaved_changes and self.current_category:
            if messagebox.askyesno("Unsaved Changes", 
                                  f"You have unsaved changes to '{self.current_category}'. Save them?"):
                self.save_details()
        
        # Check if a category is selected
        if not self.category_listbox.curselection():
            messagebox.showinfo("Info", "Please select a category to rename")
            return
            
        # Get the selected category
        selected_index = self.category_listbox.curselection()[0]
        old_name = self.category_listbox.get(selected_index)
        
        # Ask for new name
        new_name = simpledialog.askstring("Rename Category", "Enter new category name:", 
                                         parent=self, initialvalue=old_name)
        if not new_name or not new_name.strip() or new_name == old_name:
            return
            
        new_name = new_name.strip()
            
        # Check if new name already exists
        if new_name in self.categories and new_name != old_name:
            messagebox.showerror("Error", f"Category '{new_name}' already exists")
            return
            
        # Update category
        self.categories[new_name] = self.categories.pop(old_name)
        
        # Update UI
        self.populate_categories()
        
        # Find and select the renamed category
        for i in range(self.category_listbox.size()):
            if self.category_listbox.get(i) == new_name:
                self.category_listbox.selection_clear(0, tk.END)
                self.category_listbox.selection_set(i)
                self.category_listbox.see(i)
                self.current_category = new_name
                break
                
        # Update UI
        self.name_var.set(new_name)
        self.folder_var.set(self.categories[new_name].get("folder", new_name.capitalize()))
        self.abbr_var.set(self.categories[new_name].get("abbreviation", new_name.upper()[:4]))
        self.refresh_keywords()
        self.update_ui_state()
        self.has_unsaved_changes = False
                
        # Show success message
        self.status_var.set(f"Category renamed to '{new_name}'")
    
    def remove_category(self):
        """Remove the selected category"""
        # Check if a category is selected
        if not self.category_listbox.curselection():
            messagebox.showinfo("Info", "Please select a category to remove")
            return
            
        # Get the selected category
        selected_index = self.category_listbox.curselection()[0]
        category = self.category_listbox.get(selected_index)
        
        # Confirm removal
        if not messagebox.askyesno("Confirm", f"Are you sure you want to remove the category '{category}'?"):
            return
            
        # Remove category
        del self.categories[category]
        
        # Update the list
        self.populate_categories()
        
        # Reset current category
        self.current_category = None
        self.has_unsaved_changes = False
        
        # Select first category if available
        if self.category_listbox.size() > 0:
            self.category_listbox.selection_set(0)
            self.category_listbox.event_generate('<<ListboxSelect>>')
            
        # Show success message
        self.status_var.set(f"Category '{category}' removed")
    
    def auto_capitalize_folder(self):
        """Automatically generate a capitalized folder name from category name"""
        if self.current_category:
            self.folder_var.set(self.current_category.capitalize())
            self.has_unsaved_changes = True
            # Keep focus in the folder field
            self.folder_entry.focus_set()
    
    def auto_generate_abbreviation(self):
        """Automatically generate an abbreviation from category name"""
        if self.current_category:
            # Generate abbreviation - either first 4 chars uppercase or acronym
            words = self.current_category.split()
            if len(words) > 1:
                # Try to create an acronym (first letter of each word)
                abbr = "".join(word[0] for word in words if word)
                if len(abbr) < 2:  # If acronym is too short
                    abbr = self.current_category.upper()[:4]
            else:
                abbr = self.current_category.upper()[:4]
                
            self.abbr_var.set(abbr)
            self.has_unsaved_changes = True
            # Keep focus in the abbreviation field
            self.abbr_entry.focus_set()

class DateFormatDialog(tk.Toplevel):
    def __init__(self, parent, current_format, callback):
        super().__init__(parent)
        self.title("Date Format Settings")
        self.geometry("400x300")
        self.resizable(True, True)
        self.parent = parent
        self.current_format = current_format
        self.callback = callback
        
        self.setup_ui()
        
        # Make dialog modal
        self.transient(self.parent)
        self.grab_set()
        
    def setup_ui(self):
        # Main frame
        self.main_frame = ttk.Frame(self, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Format selection
        ttk.Label(self.main_frame, text="Select Date Format:").grid(row=0, column=0, pady=(0, 10), sticky=tk.W)
        
        # Radio buttons for format selection
        self.format_var = tk.StringVar(value=self.current_format)
        
        formats = [
            ("DDMMYY (e.g., 311223)", "ddmmyy"),
            ("MMDDYY (e.g., 123123)", "mmddyy"),
            ("YYMMDD (e.g., 231231)", "yymmdd")
        ]
        
        for i, (text, value) in enumerate(formats):
            ttk.Radiobutton(
                self.main_frame, 
                text=text,
                value=value,
                variable=self.format_var
            ).grid(row=i+1, column=0, pady=5, sticky=tk.W)
        
        # Example frame
        example_frame = ttk.LabelFrame(self.main_frame, text="Example")
        example_frame.grid(row=4, column=0, pady=(15, 5), sticky=(tk.W, tk.E))
        
        ttk.Label(example_frame, text="Today's date:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.example_var = tk.StringVar()
        ttk.Label(example_frame, textvariable=self.example_var).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Update example when format changes
        self.format_var.trace_add("write", lambda *args: self.update_example())
        self.update_example()
        
        # Buttons
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=5, column=0, pady=(15, 0), sticky=(tk.E))
        
        ttk.Button(button_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Save", command=self.save_format).pack(side=tk.RIGHT)
        
    def update_example(self):
        today = datetime.now()
        date_format = self.format_var.get()
        
        if date_format == "ddmmyy":
            formatted_date = today.strftime("%d%m%y")
        elif date_format == "mmddyy":
            formatted_date = today.strftime("%m%d%y")
        elif date_format == "yymmdd":
            formatted_date = today.strftime("%y%m%d")
        else:
            formatted_date = today.strftime("%d%m%y")  # Default
            
        self.example_var.set(formatted_date)
        
    def save_format(self):
        selected_format = self.format_var.get()
        self.callback(selected_format)
        self.destroy()

class PDFOrganizer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Organizer")
        self.geometry("1000x700")  # Smaller window size for 1280x900 screens
        self.minsize(800, 600)  # Set minimum window size
        
        # Define folder paths first
        self.sorted_folder = "sorted"
        self.needs_processing_folder = "needs_further_processing"
        
        # Load categories from JSON
        self.categories = self.load_categories()
        
        # Now that folder paths are defined, ensure all folders exist
        self.ensure_category_folders()
        
        # Ensure sorted folder exists
        if not os.path.exists(self.sorted_folder):
            os.makedirs(self.sorted_folder)
        
        # Ensure "needs further processing" folder exists
        if not os.path.exists(self.needs_processing_folder):
            os.makedirs(self.needs_processing_folder)
        
        # Load settings or use defaults
        self.settings = self.load_settings()
        
        # Pagination settings
        self.page_size = 100  # Number of files per page
        self.current_page = 1
        self.total_pages = 1
        self.all_pdfs = []  # Store all PDF filenames
        
        # Set up the main frame
        self.setup_ui()
        
        # Load PDFs from current directory
        self.load_all_pdfs()
        
        # Current file info
        self.current_file = None
        self.current_text = ""
        
    def load_categories(self):
        # Load categories from JSON file
        try:
            with open("categories.json", "r") as f:
                return json.load(f)
        except FileNotFoundError:
            # Create empty categories dictionary instead of default categories
            empty_categories = {}
            with open("categories.json", "w") as f:
                json.dump(empty_categories, f, indent=4)
            return empty_categories
    
    def ensure_category_folders(self):
        """Create folders for each category if they don't exist and handle special characters in paths"""
        created_folders = []
        
        # First ensure the base folders exist
        for folder in [self.sorted_folder, self.needs_processing_folder]:
            if not os.path.exists(folder):
                try:
                    os.makedirs(folder)
                    created_folders.append(folder)
                except Exception as e:
                    print(f"Error creating folder {folder}: {str(e)}")
        
        # Then create category folders
        for category, data in self.categories.items():
            folder_path = data.get("folder", category.capitalize())
            
            # Sanitize folder path by removing newlines and invalid characters
            folder_path = folder_path.replace('\n', ' ').strip()
            
            # Make sure folder path is valid
            if not folder_path:
                folder_path = category.capitalize()
                # Update the category data with the valid folder name
                self.categories[category]["folder"] = folder_path
                
            try:
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)
                    created_folders.append(folder_path)
            except Exception as e:
                # If there's an error, try to create a simplified version of the folder name
                print(f"Error creating folder '{folder_path}': {str(e)}")
                simplified_path = re.sub(r'[\\/:*?"<>|]', '_', category.capitalize())
                try:
                    if not os.path.exists(simplified_path):
                        os.makedirs(simplified_path)
                        created_folders.append(simplified_path)
                    # Update the category data with the simplified folder name
                    self.categories[category]["folder"] = simplified_path
                    print(f"Created simplified folder '{simplified_path}' instead")
                except Exception as e2:
                    print(f"Error creating simplified folder '{simplified_path}': {str(e2)}")
        
        # If any folders were created, save the updated categories
        if any(folder_path != category.capitalize() for category, data in self.categories.items() 
               for folder_path in [data.get("folder", category.capitalize())]):
            try:
                with open("categories.json", "w") as f:
                    json.dump(self.categories, f, indent=4)
                print("Updated categories.json with simplified folder names")
            except Exception as e:
                print(f"Error saving updated categories: {str(e)}")
        
        return created_folders

    def load_settings(self):
        """Load settings from JSON file"""
        default_settings = {
            "date_format": "ddmmyy"  # Default: DDMMYY
        }
        
        try:
            with open("pdf_organizer_settings.json", "r") as f:
                return json.load(f)
        except FileNotFoundError:
            # Create default settings if file doesn't exist
            with open("pdf_organizer_settings.json", "w") as f:
                json.dump(default_settings, f, indent=4)
            return default_settings

    def save_settings(self):
        """Save settings to JSON file"""
        with open("pdf_organizer_settings.json", "w") as f:
            json.dump(self.settings, f, indent=4)
    
    def setup_ui(self):
        # Create menu bar
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)
        
        # Add File menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="Refresh", command=self.refresh_pdfs)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.quit)
        
        # Add Settings menu
        self.settings_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Settings", menu=self.settings_menu)
        self.settings_menu.add_command(label="Edit Categories", command=self.edit_categories)
        self.settings_menu.add_command(label="Date Format", command=self.edit_date_format)
        
        # Create main container frame
        self.container = ttk.Frame(self)
        self.container.pack(fill=tk.BOTH, expand=True)
        
        # Configure row and column weights for resizing
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        
        # Main layout - paned window
        self.main_paned = ttk.PanedWindow(self.container, orient=tk.HORIZONTAL)
        self.main_paned.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        
        # Left panel - File list
        self.left_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.left_frame, weight=1)
        
        self.file_frame = ttk.LabelFrame(self.left_frame, text="PDF Files")
        self.file_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Create listbox with scrollbar
        self.file_list_container = ttk.Frame(self.file_frame)
        self.file_list_container.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        self.file_listbox = tk.Listbox(self.file_list_container)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        self.file_scrollbar = ttk.Scrollbar(self.file_list_container)
        self.file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Connect scrollbar to listbox
        self.file_listbox.config(yscrollcommand=self.file_scrollbar.set)
        self.file_scrollbar.config(command=self.file_listbox.yview)
        
        # Add pagination controls
        self.pagination_frame = ttk.Frame(self.file_frame)
        self.pagination_frame.pack(fill=tk.X, padx=2, pady=2)
        
        self.prev_page_btn = ttk.Button(self.pagination_frame, text="< Prev", command=self.prev_page, width=8)
        self.prev_page_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        self.page_info_var = tk.StringVar(value="Page 1 of 1")
        self.page_info_label = ttk.Label(self.pagination_frame, textvariable=self.page_info_var)
        self.page_info_label.pack(side=tk.LEFT, padx=10, pady=2, expand=True)
        
        self.next_page_btn = ttk.Button(self.pagination_frame, text="Next >", command=self.next_page, width=8)
        self.next_page_btn.pack(side=tk.RIGHT, padx=2, pady=2)
        
        # Right panel - Main content and controls
        self.right_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.right_frame, weight=3)
        
        # Configure right frame for flexible resizing
        self.right_frame.columnconfigure(0, weight=1)
        self.right_frame.rowconfigure(0, weight=3)
        self.right_frame.rowconfigure(1, weight=1)
        
        # PDF info frame
        self.pdf_info_frame = ttk.LabelFrame(self.right_frame, text="PDF Information")
        self.pdf_info_frame.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        
        # Configure pdf_info_frame for resizing
        self.pdf_info_frame.columnconfigure(0, weight=1)
        self.pdf_info_frame.rowconfigure(0, weight=1)
        self.pdf_info_frame.rowconfigure(1, weight=0)
        
        # PDF text content
        self.text_frame = ttk.Frame(self.pdf_info_frame)
        self.text_frame.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        
        self.text_frame.columnconfigure(0, weight=1)
        self.text_frame.rowconfigure(1, weight=1)
        
        self.text_label = ttk.Label(self.text_frame, text="PDF Content Preview:")
        self.text_label.grid(row=0, column=0, sticky="w", padx=2, pady=2)
        
        # Text widget with scrollbar in a frame
        self.text_container = ttk.Frame(self.text_frame)
        self.text_container.grid(row=1, column=0, sticky="nsew")
        
        self.text_box = tk.Text(self.text_container, wrap=tk.WORD, height=15)
        self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.text_scrollbar = ttk.Scrollbar(self.text_container)
        self.text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Connect scrollbar to text box
        self.text_box.config(yscrollcommand=self.text_scrollbar.set)
        self.text_scrollbar.config(command=self.text_box.yview)
        
        # Open PDF button
        self.open_button = ttk.Button(self.pdf_info_frame, text="Open PDF in Default Viewer", command=self.open_pdf_external)
        self.open_button.grid(row=1, column=0, padx=2, pady=2)
        
        # Controls frame
        self.controls_frame = ttk.LabelFrame(self.right_frame, text="File Details")
        self.controls_frame.grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        
        # Configure columns for better layout
        self.controls_frame.columnconfigure(1, weight=1)
        self.controls_frame.columnconfigure(3, weight=1)
        
        # Selected file display
        self.file_label = ttk.Label(self.controls_frame, text="Selected File:")
        self.file_label.grid(row=0, column=0, padx=2, pady=2, sticky=tk.W)
        
        self.selected_file_var = tk.StringVar()
        self.selected_file_entry = ttk.Entry(self.controls_frame, textvariable=self.selected_file_var, state="readonly")
        self.selected_file_entry.grid(row=0, column=1, columnspan=3, padx=2, pady=2, sticky=tk.W+tk.E)
        
        # Category selection
        self.category_label = ttk.Label(self.controls_frame, text="Category:")
        self.category_label.grid(row=1, column=0, padx=2, pady=2, sticky=tk.W)
        
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(self.controls_frame, textvariable=self.category_var, width=15)
        self.category_combo['values'] = list(self.categories.keys())
        self.category_combo.grid(row=1, column=1, padx=2, pady=2, sticky=tk.W)
        
        # Date field
        date_format = self.settings.get("date_format", "ddmmyy").upper()
        self.date_label = ttk.Label(self.controls_frame, text=f"Date ({date_format}):")
        self.date_label.grid(row=2, column=0, padx=2, pady=2, sticky=tk.W)
        
        self.date_var = tk.StringVar()
        self.date_entry = ttk.Entry(self.controls_frame, textvariable=self.date_var, width=8)
        self.date_entry.grid(row=2, column=1, columnspan=2, padx=2, pady=2, sticky=tk.W)
        
        # Specific naming field
        self.specific_label = ttk.Label(self.controls_frame, text="Specific:")
        self.specific_label.grid(row=3, column=0, padx=2, pady=2, sticky=tk.W)
        
        self.specific_var = tk.StringVar()
        self.specific_entry = ttk.Entry(self.controls_frame, textvariable=self.specific_var, width=15)
        self.specific_entry.grid(row=3, column=1, columnspan=2, padx=2, pady=2, sticky=tk.W+tk.E)
        
        # Preview of new filename
        self.preview_label = ttk.Label(self.controls_frame, text="New Filename:")
        self.preview_label.grid(row=4, column=0, padx=2, pady=2, sticky=tk.W)
        
        self.preview_var = tk.StringVar()
        self.preview_entry = ttk.Entry(self.controls_frame, textvariable=self.preview_var, state="readonly")
        self.preview_entry.grid(row=4, column=1, columnspan=3, padx=2, pady=2, sticky=tk.W+tk.E)
        
        # Detected category
        self.detected_label = ttk.Label(self.controls_frame, text="Detected Category:")
        self.detected_label.grid(row=5, column=0, padx=2, pady=2, sticky=tk.W)
        
        self.detected_var = tk.StringVar()
        self.detected_entry = ttk.Entry(self.controls_frame, textvariable=self.detected_var, state="readonly", width=15)
        self.detected_entry.grid(row=5, column=1, padx=2, pady=2, sticky=tk.W)
        
        # Apply detected button
        self.apply_button = ttk.Button(self.controls_frame, text="Apply", command=self.apply_detected, width=8)
        self.apply_button.grid(row=5, column=2, padx=2, pady=2, sticky=tk.W)
        
        # Detected date
        self.detected_date_label = ttk.Label(self.controls_frame, text="Detected Date:")
        self.detected_date_label.grid(row=6, column=0, padx=2, pady=2, sticky=tk.W)
        
        self.detected_date_var = tk.StringVar()
        self.detected_date_entry = ttk.Entry(self.controls_frame, textvariable=self.detected_date_var, state="readonly", width=15)
        self.detected_date_entry.grid(row=6, column=1, padx=2, pady=2, sticky=tk.W)
        
        # Apply detected date button
        self.apply_date_button = ttk.Button(self.controls_frame, text="Apply", command=self.apply_detected_date, width=8)
        self.apply_date_button.grid(row=6, column=2, padx=2, pady=2, sticky=tk.W)
        
        # Bottom row buttons
        self.button_frame = ttk.Frame(self.controls_frame)
        self.button_frame.grid(row=7, column=0, columnspan=4, sticky="ew", padx=2, pady=5)
        self.button_frame.columnconfigure(0, weight=1)  # Give full weight to the first column
        
        # Add Automate button
        self.automate_button = ttk.Button(self.button_frame, text="Auto Process All", command=self.auto_process_all, width=15)
        self.automate_button.grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        
        # Save button
        self.save_button = ttk.Button(self.button_frame, text="Save", command=self.save_file, width=8)
        self.save_button.grid(row=0, column=1, padx=5, pady=2, sticky=tk.E)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var.set("Ready")
        
        # Bind events for live preview
        self.category_var.trace_add("write", self.on_field_change)
        self.date_var.trace_add("write", self.on_field_change)
        self.specific_var.trace_add("write", self.on_field_change)
    
    def edit_categories(self):
        # Open category editor dialog
        editor = CategoryEditor(self, self.categories, self.update_categories)
        self.wait_window(editor)
    
    def update_categories(self, new_categories):
        # Update categories
        self.categories = new_categories
        
        # Save to file
        with open("categories.json", "w") as f:
            json.dump(self.categories, f, indent=4)
        
        # Update UI
        self.category_combo['values'] = list(self.categories.keys())
        
        # Ensure folders exist
        self.ensure_category_folders()
        
        # Show success message
        messagebox.showinfo("Success", "Categories updated successfully")
    
    def load_all_pdfs(self):
        """Load all PDFs from current directory into memory"""
        # Clear current list of all PDFs
        self.all_pdfs = []
        
        # Get all PDF files in current directory (excluding the sorted directory)
        for file in os.listdir('.'):
            if file.lower().endswith('.pdf'):
                self.all_pdfs.append(file)
                
        # Sort files alphabetically
        self.all_pdfs.sort()
        
        # Calculate total pages
        self.total_pages = max(1, (len(self.all_pdfs) + self.page_size - 1) // self.page_size)
        
        # Ensure current page is valid
        if self.current_page > self.total_pages:
            self.current_page = self.total_pages
            
        # Load the current page
        self.load_pdfs_page()
        
        # Update status
        self.status_var.set(f"Found {len(self.all_pdfs)} PDF files")
    
    def load_pdfs_page(self):
        """Load a specific page of PDFs into the listbox"""
        # Clear current listbox
        self.file_listbox.delete(0, tk.END)
        
        # Calculate start and end indices
        start_idx = (self.current_page - 1) * self.page_size
        end_idx = min(start_idx + self.page_size, len(self.all_pdfs))
        
        # Add PDFs for the current page to the listbox
        for i in range(start_idx, end_idx):
            self.file_listbox.insert(tk.END, self.all_pdfs[i])
            
        # Update page info
        self.page_info_var.set(f"Page {self.current_page} of {self.total_pages}")
        
        # Update button states
        self.prev_page_btn["state"] = "normal" if self.current_page > 1 else "disabled"
        self.next_page_btn["state"] = "normal" if self.current_page < self.total_pages else "disabled"
    
    def next_page(self):
        """Go to the next page of PDFs"""
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.load_pdfs_page()
    
    def prev_page(self):
        """Go to the previous page of PDFs"""
        if self.current_page > 1:
            self.current_page -= 1
            self.load_pdfs_page()
    
    def refresh_pdfs(self):
        """Refresh the PDF list"""
        # Save current page before refresh
        current_page = self.current_page
        
        # Reload all PDFs
        self.load_all_pdfs()
        
        # Try to restore previous page if possible
        if current_page <= self.total_pages:
            self.current_page = current_page
            self.load_pdfs_page()
    
    def on_file_select(self, event):
        # Get selected file
        if not self.file_listbox.curselection():
            return
        
        selected_index = self.file_listbox.curselection()[0]
        filename = self.file_listbox.get(selected_index)
        self.current_file = filename
        
        # Update UI
        self.selected_file_var.set(filename)
        
        # Process PDF
        try:
            if os.path.exists(filename):
                # Extract text from PDF
                self.current_text = self.extract_text_from_pdf(filename)
                
                # Display text in the text box
                self.text_box.delete("1.0", tk.END)
                self.text_box.insert("1.0", self.current_text[:5000] + 
                                    ("\n\n[...content truncated...]" if len(self.current_text) > 5000 else ""))
                
                # Extract date from PDF content
                date = self.extract_date_from_pdf(self.current_text)
                formatted_date = ""
                
                if date:
                    # Format the date according to the current setting
                    date_format = self.settings.get("date_format", "ddmmyy")
                    
                    if date_format == "ddmmyy":
                        formatted_date = date.strftime("%d%m%y")
                    elif date_format == "mmddyy":
                        formatted_date = date.strftime("%m%d%y")
                    elif date_format == "yymmdd":
                        formatted_date = date.strftime("%y%m%d")
                    else:
                        formatted_date = date.strftime("%d%m%y")  # Default
                        
                    self.date_var.set(formatted_date)
                    self.detected_date_var.set(formatted_date)
                else:
                    # Try from filename or default to today
                    date = self.extract_date_from_filename(filename)
                    if date:
                        # Format the date according to the current setting
                        date_format = self.settings.get("date_format", "ddmmyy")
                        
                        if date_format == "ddmmyy":
                            formatted_date = date.strftime("%d%m%y")
                        elif date_format == "mmddyy":
                            formatted_date = date.strftime("%m%d%y")
                        elif date_format == "yymmdd":
                            formatted_date = date.strftime("%y%m%d")
                        else:
                            formatted_date = date.strftime("%d%m%y")  # Default
                            
                        self.date_var.set(formatted_date)
                        self.detected_date_var.set(formatted_date)
                    else:
                        self.set_today()
                        self.detected_date_var.set("")
                
                # Detect category
                detected_category = self.detect_category(self.current_text)
                if detected_category:
                    self.detected_var.set(detected_category)
                    # Also set the category for preview
                    self.category_var.set(detected_category)
                else:
                    self.detected_var.set("")
                
                # Update status
                self.status_var.set(f"Loaded: {filename}")
            else:
                messagebox.showerror("Error", f"File not found: {filename}")
                self.status_var.set("Error loading file")
        except Exception as e:
            messagebox.showerror("Error", f"Could not process PDF: {str(e)}")
            self.status_var.set("Error processing file")
        
        # Update preview
        self.update_preview()
    
    def open_pdf_external(self):
        if not self.current_file or not os.path.exists(self.current_file):
            messagebox.showinfo("Info", "No PDF file selected")
            return
            
        # Open PDF in system default viewer
        try:
            if platform.system() == 'Darwin':  # macOS
                subprocess.call(('open', self.current_file))
            elif platform.system() == 'Windows':  # Windows
                os.startfile(os.path.abspath(self.current_file))
            else:  # Linux variants
                subprocess.call(('xdg-open', self.current_file))
                
            self.status_var.set(f"Opened {self.current_file} in external viewer")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF: {str(e)}")
    
    def extract_text_from_pdf(self, filename):
        try:
            with open(filename, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                # Extract text from first few pages (max 5)
                num_pages = min(len(reader.pages), 5)
                for i in range(num_pages):
                    page_text = reader.pages[i].extract_text()
                    if page_text:
                        text += page_text + "\n"
                return text
        except Exception as e:
            print(f"Error extracting text: {str(e)}")
            return ""
    
    def extract_date_from_pdf(self, text):
        # Clean up text - remove extra spaces and normalize
        clean_text = re.sub(r'\s+', ' ', text)
        
        # Look for common OCR errors in years and fix them
        # Fix cases like "202 5" or "2 023" to "2025" or "2023"
        clean_text = re.sub(r'(\b20\d{1,2})\s+(\d{1})\b', r'\1\2', clean_text)  # Fix split years like "202 5"
        clean_text = re.sub(r'(\b2)\s+(\d{3})\b', r'\1\2', clean_text)  # Fix split years like "2 023"

        # List of month names and abbreviations for pattern matching
        months = r'(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)'
        
        # Look for date patterns in the text - expanded with more patterns
        date_patterns = [
            # Standard formats with separators
            r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})',  # DD/MM/YYYY or MM/DD/YYYY
            r'(\d{2,4})[/.-](\d{1,2})[/.-](\d{1,2})',  # YYYY/MM/DD
            
            # Text formats
            rf'(\d{{1,2}})(?:st|nd|rd|th)?\s+(?:of\s+)?({months})[,\s]+(\d{{2,4}})',  # DD Month YYYY
            rf'({months})\s+(\d{{1,2}})(?:st|nd|rd|th)?[,\s]+(\d{{2,4}})',  # Month DD, YYYY
            rf'({months})\s+(\d{{1,2}})(?:st|nd|rd|th)?[,\s]+(20\d{{2}})',  # Month DD 20XX specific for recent years
            
            # ISO format
            r'(\d{4})-(\d{2})-(\d{2})',  # YYYY-MM-DD
            
            # Formats with text month and no separators
            rf'(\d{{1,2}})\s+({months})\s+(\d{{4}})',  # DD Month YYYY without commas
            
            # Special fixes for OCR errors
            r'(\d{1,2})[/.-](\d{1,2})[/.-]\s*(\d{2,4})',  # Handle space before year
            r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{2})(\d{2})',  # Split year like 20 23
        ]
        
        # First try dateutil parser on the whole text - handles many formats automatically
        try:
            date = dateutil.parser.parse(clean_text, fuzzy=True)
            # Verify the date is reasonable (between 1900 and 2100)
            if 1900 <= date.year <= 2100:
                return date
        except:
            pass  # If it fails, continue with pattern matching
        
        # Now try our custom patterns
        for pattern in date_patterns:
            matches = re.findall(pattern, clean_text)
            if matches:
                for match in matches:
                    try:
                        # Turn tuple into a string for dateutil to parse
                        date_str = ' '.join(str(part) for part in match if part)
                        
                        # Use parse to handle various date formats
                        date = dateutil.parser.parse(date_str, fuzzy=True)
                        
                        # Ensure date is valid (year >= 1900 to avoid Windows formatting issues)
                        if 1900 <= date.year <= 2100:  # Add reasonable upper bound
                            return date
                    except:
                        continue
        
        # If we get here, try more aggressive search for just a year
        year_matches = re.findall(r'\b(19\d{2}|20\d{2})\b', clean_text)
        if year_matches:
            try:
                # If we just found a year, use today's month and day with that year
                today = datetime.now()
                year = int(year_matches[0])
                if 1900 <= year <= 2100:
                    return datetime(year, today.month, today.day)
            except:
                pass
        
        return None
    
    def extract_date_from_filename(self, filename):
        # Clean filename - remove potential OCR artifacts
        clean_filename = re.sub(r'\s+', '', filename)  # Remove all spaces
        
        # Try to find date pattern in filename - expanded patterns
        date_patterns = [
            r'(\d{2})(\d{2})(\d{2,4})',  # DDMMYY or DDMMYYYY
            r'(\d{2,4})(\d{2})(\d{2})',  # YYYYMMDD or YYMMDD
            r'(\d{2})[_-](\d{2})[_-](\d{2,4})',  # DD-MM-YY or DD_MM_YYYY
            r'(\d{4})[_-](\d{2})[_-](\d{2})',  # YYYY-MM-DD ISO format
            r'(\d{2})[_-](\d{2})[_-](\d{2})'    # YY-MM-DD
        ]
        
        # First try clean filename
        for pattern in date_patterns:
            matches = re.findall(pattern, clean_filename)
            found_date = self._process_date_matches(matches)
            if found_date:
                return found_date
        
        # If that fails, try the original filename
        for pattern in date_patterns:
            matches = re.findall(pattern, filename)
            found_date = self._process_date_matches(matches)
            if found_date:
                return found_date
                
        return None
    
    def _process_date_matches(self, matches):
        """Helper to process date matches from a pattern"""
        if not matches:
            return None
            
        for match in matches:
            try:
                # Try to determine date format
                part1, part2, part3 = match[0], match[1], match[2]
                
                # Try to determine if it's YYYYMMDD, DDMMYYYY, etc.
                if len(part1) == 4:  # First part is 4 digits, likely YYYY
                    year = int(part1)
                    month = int(part2)
                    day = int(part3)
                elif len(part3) == 4:  # Last part is 4 digits, likely YYYY
                    day = int(part1)
                    month = int(part2)
                    year = int(part3)
                else:
                    # Handle 2-digit years
                    if int(part1) <= 31 and int(part2) <= 12:
                        # Likely DD-MM-YY
                        day = int(part1)
                        month = int(part2)
                        year = 2000 + int(part3) if int(part3) < 50 else 1900 + int(part3)
                    elif int(part1) <= 12 and int(part2) <= 31:
                        # Likely MM-DD-YY
                        month = int(part1)
                        day = int(part2)
                        year = 2000 + int(part3) if int(part3) < 50 else 1900 + int(part3)
                    else:
                        # Try YY-MM-DD
                        year = 2000 + int(part1) if int(part1) < 50 else 1900 + int(part1)
                        month = int(part2)
                        day = int(part3)
                
                # Validate month and day
                if not (1 <= month <= 12 and 1 <= day <= 31):
                    continue
                
                # Ensure year is valid for Windows (>= 1900)
                if year < 1900 or year > 2100:
                    continue
                    
                # Create date object
                return datetime(year, month, day)
            except:
                continue
        
        return None
    
    def detect_category(self, text):
        # Convert text to lowercase for case-insensitive matching
        text_lower = text.lower()
        
        # Normalize spaces in text (replace multiple spaces with single space)
        normalized_text = re.sub(r'\s+', ' ', text_lower)
        
        best_match = None
        max_matches = 0
        
        # Check each category's keywords
        for category, data in self.categories.items():
            keywords = data.get("keywords", [])
            if not keywords:
                continue
                
            matches = 0
            for keyword in keywords:
                # Normalize spaces in keyword too
                normalized_keyword = re.sub(r'\s+', ' ', keyword.lower())
                
                # Try both original and normalized matching
                if normalized_keyword in normalized_text or keyword.lower() in text_lower:
                    matches += 1
            
            # Update if this category has more matching keywords
            if matches > max_matches:
                max_matches = matches
                best_match = category
        
        return best_match
    
    def apply_detected(self):
        detected = self.detected_var.get()
        if detected:
            self.category_var.set(detected)
            self.update_preview()
    
    def apply_detected_date(self):
        detected_date = self.detected_date_var.get()
        if detected_date:
            self.date_var.set(detected_date)
            self.update_preview()
    
    def set_today(self):
        # Set date field to today's date using the configured format
        date_format = self.settings.get("date_format", "ddmmyy")
        
        today = datetime.now()
        try:
            if date_format == "ddmmyy":
                formatted_date = today.strftime("%d%m%y")
            elif date_format == "mmddyy":
                formatted_date = today.strftime("%m%d%y")
            elif date_format == "yymmdd":
                formatted_date = today.strftime("%y%m%d")
            else:
                formatted_date = today.strftime("%d%m%y")  # Default
                
            self.date_var.set(formatted_date)
        except ValueError as e:
            messagebox.showerror("Date Error", f"Error formatting date: {str(e)}")
            self.date_var.set("")
    
    def on_field_change(self, *args):
        # Update preview when any field changes
        self.update_preview()
    
    def update_preview(self):
        if not self.current_file:
            self.preview_var.set("")
            return
            
        category = self.category_var.get()
        date_str = self.date_var.get()
        specific = self.specific_var.get()
        
        if not category:
            self.preview_var.set("")
            return
            
        # Get abbreviation for category
        abbr = self.categories.get(category, {}).get("abbreviation", category.upper()[:4])
        
        # Create preview filename
        if specific:
            new_filename = f"{date_str}_{abbr}_{specific}.pdf"
        else:
            new_filename = f"{date_str}_{abbr}.pdf"
            
        self.preview_var.set(new_filename)
    
    def save_file(self):
        if not self.current_file:
            messagebox.showinfo("Info", "No file selected")
            return
            
        category = self.category_var.get()
        if not category:
            messagebox.showinfo("Info", "Please select a category")
            return
            
        date_str = self.date_var.get()
        if not re.match(r'^\d{6}$', date_str):
            messagebox.showinfo("Info", "Date must be in 6-digit format (e.g., 311223)")
            return
            
        # Get abbreviation for category
        abbr = self.categories.get(category, {}).get("abbreviation", category.upper()[:4])
        
        # Create new filename - always use the simple format
        new_filename = f"{date_str}_{abbr}.pdf"
            
        # Get destination folder
        folder = self.categories.get(category, {}).get("folder", category.capitalize())
        
        # Create full destination path
        destination = os.path.join(folder, new_filename)
        
        # Check if file already exists
        if os.path.exists(destination):
            # Extract base name (without extension) and extension
            basename, ext = os.path.splitext(new_filename)
            
            # Find next available number
            counter = 1
            while os.path.exists(os.path.join(folder, f"{basename} ({counter}){ext}")):
                counter += 1
                
            # Update destination with numbered suffix
            destination = os.path.join(folder, f"{basename} ({counter}){ext}")
        
        # Get current selection index for later use
        current_index = -1
        if self.file_listbox.curselection():
            current_index = self.file_listbox.curselection()[0]
        
        # Copy file to destination
        try:
            # Copy file to category folder with new name
            shutil.copy2(self.current_file, destination)
            
            # Move original file to sorted folder
            sorted_destination = os.path.join(self.sorted_folder, os.path.basename(self.current_file))
            
            # Check if a file with the same name already exists in the sorted folder
            if os.path.exists(sorted_destination):
                # Add a timestamp to make filename unique
                filename, ext = os.path.splitext(os.path.basename(self.current_file))
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                sorted_destination = os.path.join(self.sorted_folder, f"{filename}_{timestamp}{ext}")
            
            # Move the original file to sorted folder
            shutil.move(self.current_file, sorted_destination)
            
            # Remove the file from the all_pdfs list
            if self.current_file in self.all_pdfs:
                self.all_pdfs.remove(self.current_file)
                
            # Update total pages
            self.total_pages = max(1, (len(self.all_pdfs) + self.page_size - 1) // self.page_size)
            
            # Ensure current page is valid
            if self.current_page > self.total_pages:
                self.current_page = self.total_pages
            
            self.status_var.set(f"File saved as {destination} and original moved to {self.sorted_folder} folder")
            messagebox.showinfo("Success", f"File saved as {destination}\nOriginal moved to {self.sorted_folder} folder")
            
            # Clear current file since it's been moved
            self.current_file = None
            
            # Reload PDFs for the current page
            self.load_pdfs_page()
            
            # Select the next item in the list
            new_size = self.file_listbox.size()
            if new_size > 0:
                # If there are still items in the list
                if current_index >= new_size:
                    # If the previous selection was the last item or beyond, select the last item
                    self.file_listbox.selection_set(new_size - 1)
                    self.file_listbox.see(new_size - 1)
                else:
                    # Select the item at the same index as before
                    self.file_listbox.selection_set(current_index)
                    self.file_listbox.see(current_index)
                
                # Trigger the file selection event to load the PDF
                self.file_listbox.event_generate('<<ListboxSelect>>')
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not save/move file: {str(e)}")

    def edit_date_format(self):
        """Open dialog to edit date format"""
        dialog = DateFormatDialog(self, self.settings.get("date_format", "ddmmyy"), self.update_date_format)
        self.wait_window(dialog)
        
    def update_date_format(self, new_format):
        """Update date format setting"""
        self.settings["date_format"] = new_format
        self.save_settings()
        
        # Update date label to show new format
        date_format = new_format.upper()
        self.date_label.config(text=f"Date ({date_format}):")
        
        messagebox.showinfo("Success", "Date format updated successfully")
        
        # Update current date field if a file is selected
        if self.current_file:
            self.set_today()  # This will update using the new format

    def auto_process_all(self):
        """Automatically process all PDFs in the current directory"""
        if not self.all_pdfs:
            messagebox.showinfo("Info", "No PDF files found to process")
            return
            
        # Verify all needed folders exist before starting
        if not self.verify_folders_before_processing():
            messagebox.showerror("Error", "Unable to create necessary folders. Please check folder permissions and try again.")
            return
            
        # Ask for confirmation
        if not messagebox.askyesno("Confirm", f"This will automatically process {len(self.all_pdfs)} PDF files.\nContinue?"):
            return
            
        # First pass: Analyze all PDFs to identify ones that need manual intervention
        manual_processing_needed = []
        analysis_results = {}
        
        # Create progress dialog for analysis
        analysis_window = tk.Toplevel(self)
        analysis_window.title("Analyzing PDFs")
        analysis_window.geometry("400x150")
        analysis_window.resizable(False, False)
        analysis_window.transient(self)
        analysis_window.grab_set()
        
        # Analysis progress frame
        analysis_frame = ttk.Frame(analysis_window, padding=10)
        analysis_frame.pack(fill=tk.BOTH, expand=True)
        
        # Status label
        analysis_status_var = tk.StringVar(value="Analyzing files...")
        ttk.Label(analysis_frame, textvariable=analysis_status_var).pack(pady=(0, 10))
        
        # Current file label
        current_file_var = tk.StringVar(value="")
        ttk.Label(analysis_frame, textvariable=current_file_var).pack(pady=(0, 10))
        
        # Progress bar
        analysis_progress_var = tk.DoubleVar(value=0.0)
        analysis_progress_bar = ttk.Progressbar(analysis_frame, variable=analysis_progress_var, maximum=len(self.all_pdfs))
        analysis_progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Update UI
        analysis_window.update()
        
        # Analyze each PDF
        for i, pdf_file in enumerate(self.all_pdfs):
            # Update progress UI
            analysis_progress_var.set(i + 1)
            current_file_var.set(f"Analyzing: {pdf_file}")
            analysis_window.update()
            
            try:
                if os.path.exists(pdf_file):
                    # Extract text from PDF
                    pdf_text = self.extract_text_from_pdf(pdf_file)
                    
                    # Extract date
                    detected_date = self.extract_date_from_pdf(pdf_text)
                    if not detected_date:
                        detected_date = self.extract_date_from_filename(pdf_file)
                    
                    # Detect category
                    detected_category, confidence = self.detect_category_with_confidence(pdf_text)
                    
                    # Store results for later use
                    analysis_results[pdf_file] = {
                        "text": pdf_text,
                        "date": detected_date,
                        "category": detected_category,
                        "confidence": confidence
                    }
                    
                    # If missing date or category, mark for manual processing
                    if not detected_date or not detected_category or confidence < 1:
                        manual_processing_needed.append(pdf_file)
            except Exception as e:
                print(f"Error analyzing {pdf_file}: {str(e)}")
                manual_processing_needed.append(pdf_file)
        
        # Close analysis window
        analysis_window.destroy()
        
        # If files need manual processing, show dialog
        if manual_processing_needed:
            manual_process_result = messagebox.askyesno(
                "Manual Processing Required",
                f"{len(manual_processing_needed)} files require manual processing.\n\n"
                "Would you like to process these files now?\n\n"
                "If you select 'No', these files will be moved to the 'needs_further_processing' folder."
            )
            
            if manual_process_result:
                # Handle manual processing
                for pdf_file in manual_processing_needed:
                    # Create dialog for manual processing
                    manual_window = tk.Toplevel(self)
                    manual_window.title(f"Manual Processing: {pdf_file}")
                    manual_window.geometry("600x500")
                    manual_window.resizable(True, True)
                    manual_window.transient(self)
                    manual_window.grab_set()
                    
                    # File info frame
                    info_frame = ttk.Frame(manual_window, padding=10)
                    info_frame.pack(fill=tk.BOTH, expand=True)
                    
                    # Show file details
                    ttk.Label(info_frame, text=f"File: {pdf_file}", font=("", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
                    
                    # PDF content preview
                    ttk.Label(info_frame, text="PDF Content Preview:").pack(anchor=tk.W)
                    
                    # Text widget with scrollbar for content
                    text_frame = ttk.Frame(info_frame)
                    text_frame.pack(fill=tk.BOTH, expand=True)
                    
                    text_box = tk.Text(text_frame, wrap=tk.WORD, height=10)
                    text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                    
                    text_scrollbar = ttk.Scrollbar(text_frame)
                    text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                    
                    # Connect scrollbar to text box
                    text_box.config(yscrollcommand=text_scrollbar.set)
                    text_scrollbar.config(command=text_box.yview)
                    
                    # Insert content
                    pdf_data = analysis_results.get(pdf_file, {})
                    pdf_text = pdf_data.get("text", "")
                    text_box.insert("1.0", pdf_text[:2000] + 
                                   ("\n\n[...content truncated...]" if len(pdf_text) > 2000 else ""))
                    text_box.config(state="disabled")  # Make read-only
                    
                    # Category selection
                    category_frame = ttk.Frame(info_frame)
                    category_frame.pack(fill=tk.X, pady=(0, 5))
                    
                    ttk.Label(category_frame, text="Category:").pack(side=tk.LEFT)
                    
                    category_var = tk.StringVar()
                    if pdf_data.get("category"):
                        category_var.set(pdf_data["category"])
                        
                    category_combo = ttk.Combobox(category_frame, textvariable=category_var, width=20)
                    category_combo['values'] = list(self.categories.keys())
                    category_combo.pack(side=tk.LEFT, padx=5)
                    
                    # Date selection
                    date_frame = ttk.Frame(info_frame)
                    date_frame.pack(fill=tk.X, pady=(0, 5))
                    
                    date_format = self.settings.get("date_format", "ddmmyy")
                    date_format_display = date_format.upper()
                    
                    ttk.Label(date_frame, text=f"Date ({date_format_display}):").pack(side=tk.LEFT)
                    
                    date_var = tk.StringVar()
                    # Format the date if available
                    if pdf_data.get("date"):
                        detected_date = pdf_data["date"]
                        if date_format == "ddmmyy":
                            formatted_date = detected_date.strftime("%d%m%y")
                        elif date_format == "mmddyy":
                            formatted_date = detected_date.strftime("%m%d%y")
                        elif date_format == "yymmdd":
                            formatted_date = detected_date.strftime("%y%m%d")
                        else:
                            formatted_date = detected_date.strftime("%d%m%y")  # Default
                        date_var.set(formatted_date)
                    
                    date_entry = ttk.Entry(date_frame, textvariable=date_var, width=8)
                    date_entry.pack(side=tk.LEFT, padx=5)
                    
                    # Store the result
                    result_data = {
                        "processed": False,
                        "category": None,
                        "date": None
                    }
                    
                    # Process function
                    def process_file():
                        # Validate inputs
                        if not category_var.get():
                            messagebox.showinfo("Info", "Please select a category")
                            return
                            
                        if not re.match(r'^\d{6}$', date_var.get()):
                            messagebox.showinfo("Info", "Date must be in 6-digit format (e.g., 311223)")
                            return
                            
                        # Store result
                        result_data["processed"] = True
                        result_data["category"] = category_var.get()
                        result_data["date"] = date_var.get()
                        
                        # Close window
                        manual_window.destroy()
                    
                    # Skip function
                    def skip_file():
                        manual_window.destroy()
                    
                    # Buttons
                    buttons_frame = ttk.Frame(info_frame)
                    buttons_frame.pack(fill=tk.X, pady=(10, 0))
                    
                    process_button = ttk.Button(buttons_frame, text="Process File", command=process_file)
                    process_button.pack(side=tk.RIGHT, padx=5)
                    
                    skip_button = ttk.Button(buttons_frame, text="Skip File", command=skip_file)
                    skip_button.pack(side=tk.RIGHT, padx=5)
                    
                    # Wait for window to close
                    self.wait_window(manual_window)
                    
                    # Update analysis results with manual input
                    if result_data["processed"]:
                        analysis_results[pdf_file]["manual_category"] = result_data["category"]
                        analysis_results[pdf_file]["manual_date"] = result_data["date"]
                    else:
                        # Mark for further processing folder if skipped
                        analysis_results[pdf_file]["skip"] = True
        
        # Now process all files
        self.process_analyzed_files(analysis_results)

    def verify_folders_before_processing(self):
        """Ensure all needed folders exist before processing files"""
        # Check if needs_processing_folder exists
        if not os.path.exists(self.needs_processing_folder):
            try:
                os.makedirs(self.needs_processing_folder)
                print(f"Created missing folder: {self.needs_processing_folder}")
            except Exception as e:
                print(f"Error creating folder {self.needs_processing_folder}: {str(e)}")
                return False
        
        # Check if sorted_folder exists
        if not os.path.exists(self.sorted_folder):
            try:
                os.makedirs(self.sorted_folder)
                print(f"Created missing folder: {self.sorted_folder}")
            except Exception as e:
                print(f"Error creating folder {self.sorted_folder}: {str(e)}")
                return False
        
        # Verify all category folders exist
        for category, data in self.categories.items():
            folder_path = data.get("folder", category.capitalize())
            if not os.path.exists(folder_path):
                try:
                    os.makedirs(folder_path)
                    print(f"Created missing folder: {folder_path}")
                except Exception as e:
                    # Try to create a simplified version
                    print(f"Error creating folder '{folder_path}': {str(e)}")
                    simplified_path = re.sub(r'[\\/:*?"<>|]', '_', category.capitalize())
                    try:
                        if not os.path.exists(simplified_path):
                            os.makedirs(simplified_path)
                        # Update the category with the simplified path
                        self.categories[category]["folder"] = simplified_path
                        print(f"Created simplified folder '{simplified_path}' instead")
                    except Exception as e2:
                        print(f"Error creating simplified folder '{simplified_path}': {str(e2)}")
                        return False
        
        # If we got here, all folders exist or were created
        return True

    def process_analyzed_files(self, analysis_results):
        """Process files based on analysis results"""
        if not analysis_results:
            return
        
        # Verify all needed folders exist before starting
        if not self.verify_folders_before_processing():
            messagebox.showerror("Error", "Unable to create necessary folders. Please check folder permissions and try again.")
            return
        
        # Create progress dialog
        progress_window = tk.Toplevel(self)
        progress_window.title("Processing PDFs")
        progress_window.geometry("400x150")
        progress_window.resizable(False, False)
        progress_window.transient(self)
        progress_window.grab_set()
        
        # Progress frame
        progress_frame = ttk.Frame(progress_window, padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        # Status label
        status_var = tk.StringVar(value="Processing files...")
        status_label = ttk.Label(progress_frame, textvariable=status_var)
        status_label.pack(pady=(0, 10))
        
        # Current file label
        current_file_var = tk.StringVar(value="")
        current_file_label = ttk.Label(progress_frame, textvariable=current_file_var)
        current_file_label.pack(pady=(0, 10))
        
        # Progress bar
        progress_var = tk.DoubleVar(value=0.0)
        progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=len(analysis_results))
        progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Results variables for the summary
        processed_count = 0
        categorized_count = 0
        needs_processing_count = 0
        duplicate_count = 0
        
        # Track processed date+category combinations to avoid duplicates
        processed_combinations = {}
        
        # Create a detailed log of what happened to each file
        detailed_log = []
        
        # Process files
        progress_window.update()
        
        # Process each file
        for i, (pdf_file, data) in enumerate(analysis_results.items()):
            # Update progress UI
            progress_var.set(i + 1)
            current_file_var.set(f"Processing: {pdf_file}")
            progress_window.update()
            
            try:
                if os.path.exists(pdf_file):
                    # Determine if we should use manual or auto-detected values
                    use_manual = "manual_category" in data
                    
                    # Get category and date
                    if use_manual:
                        category = data["manual_category"]
                        date_str = data["manual_date"]
                    else:
                        # Skip files marked to skip or without sufficient info
                        if data.get("skip") or not data.get("category") or not data.get("date") or data.get("confidence", 0) < 1:
                            # Move to "needs further processing" folder
                            needs_processing_destination = os.path.join(self.needs_processing_folder, os.path.basename(pdf_file))
                            
                            # Handle filename collision
                            if os.path.exists(needs_processing_destination):
                                filename, ext = os.path.splitext(os.path.basename(pdf_file))
                                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                                needs_processing_destination = os.path.join(self.needs_processing_folder, f"{filename}_{timestamp}{ext}")
                            
                            # Move file
                            shutil.move(pdf_file, needs_processing_destination)
                            if pdf_file in self.all_pdfs:
                                self.all_pdfs.remove(pdf_file)
                            needs_processing_count += 1
                            processed_count += 1
                            
                            # Add to log
                            detailed_log.append(f"{pdf_file} → Needs further processing")
                            continue
                        
                        category = data["category"]
                        
                        # Format the date
                        detected_date = data["date"]
                        date_format = self.settings.get("date_format", "ddmmyy")
                        
                        if date_format == "ddmmyy":
                            date_str = detected_date.strftime("%d%m%y")
                        elif date_format == "mmddyy":
                            date_str = detected_date.strftime("%m%d%y")
                        elif date_format == "yymmdd":
                            date_str = detected_date.strftime("%y%m%d")
                        else:
                            date_str = detected_date.strftime("%d%m%y")  # Default
                    
                    # Get abbreviation for category
                    abbr = self.categories.get(category, {}).get("abbreviation", category.upper()[:4])
                    
                    # Generate a combination key for duplicate checking
                    combination_key = f"{date_str}_{abbr}"
                    
                    # Get destination folder
                    folder = self.categories.get(category, {}).get("folder", category.capitalize())
                    
                    # Create new filename (simple format)
                    new_filename = f"{date_str}_{abbr}.pdf"
                    
                    # Create full destination path
                    destination = os.path.join(folder, new_filename)
                    
                    # Check if this file would be a duplicate (already processed in this batch)
                    is_duplicate = combination_key in processed_combinations
                    
                    # Handle filename collision with numbered suffixes (1), (2), etc.
                    if os.path.exists(destination) or is_duplicate:
                        # Extract base name (without extension) and extension
                        basename, ext = os.path.splitext(new_filename)
                        
                        # Find next available number
                        counter = 1
                        while os.path.exists(os.path.join(folder, f"{basename} ({counter}){ext}")):
                            counter += 1
                            
                        # Update destination with numbered suffix
                        destination = os.path.join(folder, f"{basename} ({counter}){ext}")
                        
                        # Update duplicate count if it was a duplicate in this batch
                        if is_duplicate:
                            duplicate_count += 1
                    
                    # Copy to category folder
                    shutil.copy2(pdf_file, destination)
                    
                    # Add to processed combinations
                    processed_combinations[combination_key] = True
                    
                    # Move original to sorted folder
                    sorted_destination = os.path.join(self.sorted_folder, os.path.basename(pdf_file))
                    
                    # Handle filename collision in sorted folder
                    if os.path.exists(sorted_destination):
                        filename, ext = os.path.splitext(os.path.basename(pdf_file))
                        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                        sorted_destination = os.path.join(self.sorted_folder, f"{filename}_{timestamp}{ext}")
                    
                    # Move file
                    shutil.move(pdf_file, sorted_destination)
                    if pdf_file in self.all_pdfs:
                        self.all_pdfs.remove(pdf_file)
                    categorized_count += 1
                    processed_count += 1
                    
                    # Add to log
                    dest_filename = os.path.basename(destination)
                    log_entry = f"{pdf_file} → {folder}/{dest_filename}"
                    if is_duplicate:
                        log_entry += " (duplicate)"
                    detailed_log.append(log_entry)
                    
            except Exception as e:
                print(f"Error processing {pdf_file}: {str(e)}")
                # Move to needs_processing on error
                try:
                    needs_processing_destination = os.path.join(self.needs_processing_folder, os.path.basename(pdf_file))
                    
                    # Handle filename collision
                    if os.path.exists(needs_processing_destination):
                        filename, ext = os.path.splitext(os.path.basename(pdf_file))
                        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                        needs_processing_destination = os.path.join(self.needs_processing_folder, f"{filename}_{timestamp}{ext}")
                    
                    # Move file
                    if os.path.exists(pdf_file):
                        shutil.move(pdf_file, needs_processing_destination)
                        if pdf_file in self.all_pdfs:
                            self.all_pdfs.remove(pdf_file)
                        needs_processing_count += 1
                        processed_count += 1
                        
                        # Add to log
                        detailed_log.append(f"{pdf_file} → Needs further processing (error: {str(e)})")
                        
                except Exception as move_error:
                    print(f"Error moving file to needs processing: {str(move_error)}")
                    detailed_log.append(f"{pdf_file} → ERROR: {str(e)}, then {str(move_error)}")
        
        # Close progress window
        progress_window.destroy()
        
        # Reload PDFs list
        self.load_all_pdfs()
        
        # Create a results log window
        self.show_processing_log(processed_count, categorized_count, duplicate_count, 
                                needs_processing_count, detailed_log)
    
    def show_processing_log(self, processed_count, categorized_count, duplicate_count, 
                           needs_processing_count, detailed_log):
        """Display a dialog with processing results and copyable log"""
        log_window = tk.Toplevel(self)
        log_window.title("Processing Results")
        log_window.geometry("800x500")
        log_window.resizable(True, True)
        log_window.transient(self)
        
        # Create main frame
        main_frame = ttk.Frame(log_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add summary at the top
        summary_frame = ttk.LabelFrame(main_frame, text="Processing Summary", padding="10")
        summary_frame.pack(fill=tk.X, pady=(0, 10))
        
        summary_text = (
            f"Processed: {processed_count} files\n"
            f"Successfully categorized: {categorized_count}\n"
            f"Duplicate files: {duplicate_count}\n"
            f"Needs further processing: {needs_processing_count}"
        )
        
        summary_label = ttk.Label(summary_frame, text=summary_text)
        summary_label.pack(anchor=tk.W)
        
        # Create detailed log area
        log_frame = ttk.LabelFrame(main_frame, text="Detailed Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Text widget with scrollbar for log content
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        log_text = tk.Text(text_frame, wrap=tk.NONE, font=("Courier", 10))
        log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Vertical scrollbar
        v_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=log_text.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        log_text.config(yscrollcommand=v_scrollbar.set)
        
        # Horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(log_frame, orient="horizontal", command=log_text.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        log_text.config(xscrollcommand=h_scrollbar.set)
        
        # Insert log content
        log_content = "\n".join(detailed_log)
        log_text.insert("1.0", log_content)
        log_text.config(state="disabled")  # Make read-only
        
        # Add button bar at the bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Copy button
        copy_button = ttk.Button(button_frame, text="Copy to Clipboard", 
                               command=lambda: self.copy_log_to_clipboard(log_content))
        copy_button.pack(side=tk.LEFT, padx=5)
        
        # Save button
        save_button = ttk.Button(button_frame, text="Save Log File", 
                               command=lambda: self.save_log_to_file(log_content))
        save_button.pack(side=tk.LEFT, padx=5)
        
        # Close button
        close_button = ttk.Button(button_frame, text="Close", command=log_window.destroy)
        close_button.pack(side=tk.RIGHT, padx=5)

    def copy_log_to_clipboard(self, log_content):
        """Copy log content to clipboard"""
        self.clipboard_clear()
        self.clipboard_append(log_content)
        self.status_var.set("Log copied to clipboard")

    def save_log_to_file(self, log_content):
        """Save log content to a file"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"pdf_organizer_log_{timestamp}.txt"
        
        file_path = filedialog.asksaveasfilename(
            initialfile=default_filename,
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, "w") as f:
                    # Add a header with timestamp
                    header = f"PDF Organizer Processing Log - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f.write(header)
                    f.write("=" * len(header) + "\n\n")
                    f.write(log_content)
                    
                self.status_var.set(f"Log saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save log file: {str(e)}")
    
    def detect_category_with_confidence(self, text):
        """Detect category from text and return the confidence level"""
        # Convert text to lowercase for case-insensitive matching
        text_lower = text.lower()
        
        # Normalize spaces in text (replace multiple spaces with single space)
        normalized_text = re.sub(r'\s+', ' ', text_lower)
        
        best_match = None
        max_matches = 0
        
        # Check each category's keywords
        for category, data in self.categories.items():
            keywords = data.get("keywords", [])
            if not keywords:
                continue
                
            matches = 0
            for keyword in keywords:
                # Normalize spaces in keyword too
                normalized_keyword = re.sub(r'\s+', ' ', keyword.lower())
                
                # Try both original and normalized matching
                if normalized_keyword in normalized_text or keyword.lower() in text_lower:
                    matches += 1
            
            # Update if this category has more matching keywords
            if matches > max_matches:
                max_matches = matches
                best_match = category
        
        return best_match, max_matches

if __name__ == "__main__":
    app = PDFOrganizer()
    app.mainloop() 
