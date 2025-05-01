import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText
import customtkinter as ctk  # More modern UI components
from inventory_checker import InventoryChecker
from datetime import datetime
import io
import contextlib
import base64
from PIL import Image, ImageTk
from io import BytesIO

# App icon as base64 string (wine bottle icon)
APP_ICON_BASE64 = """
iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAACXBIWXMAAAsTAAALEwEAmpwYAAANTklEQVR4nO2ae3AU1R3HPzM7u5s3rwDhKY8QEiIPeQg+SEXaoFgQq7VVQLTWIihS+rDaVkextmM7Th+2HUWKreADbV+IFQQftSoRRFEJCSYkEQIECCHZ7G52d+/uXf+IYbO72ZBAYqbj9zMz2b33e8/fPee8fueec5dYJu0lXOUQQhM8gpSrCTwJVXqpN3Qer9xfuFxNkD3dQT2B89OA4aEJQsotQsjv4jC+I9DfklJ+EZqQnzVFOsLlGgZHdAfQq8hfnD04fWX+6FF3jhl2Xd6YrEzjIJMxDSEEUggJAillpJtHlJUeRVnhVZTXPygu+WJn4e7SVLOp3g3XT8rLdlVV1fxrw6bq+x57ut7hcBD47rJAYDswu/2Hpk/rrR5PB/QpA9SGjvO9wy5/bfHcWXl5w7Mxq6CEMyGO2xf7YMrYdMfRjYDN5/dPe+rFV7aUHj9h9vnDJCCvGDOqeOcu+y+eeqG60+HofEDaYUBSdoA8gMB9Y0FBwYoF82bnZmekCUMwcEJ0sUIIbDYb0wryGDd6RObRsuNLGpzOgP8CDAiALwFnWW1/de/W1mjOk+0DQdJzIdkg8Y4Q3XSzAnjx63+9cL5eRnpaQLV+XZD1RVhGApMJhNTJoyj0e3GrihzkFFNPgngOL71P77y+iqXqz6sE9QXpAQ6ApJmQKTdERTl3fvvvWfKghlTM0vKHYL2bfcmwnMJi8nExHE5HD5W4m9yuycAlwBSagZPAbLm/cfWV5QeL8PnC5oQ2F8xgwTJ6oRw56WUK2+75eaH5s+enm6yWIRmdJJMOEXPJyGE8Pt8zHrwp5wur0gFcoFpITfeFJxOvvbmO9tq6+qQUoCiQmBQZOB0qiBnFhcfL31g5Yo3RuVkG0RXjr/a8wkmk4l77piD2Wq9BNS0l0tACXA5OVX2r23bdyKlH6MxWnYbgxQDAtUblD/06COL7p4+JV0GSlqdoVeNbzveZDYzf/YMVFU1Am8HngWAGqDU7fXu2LpjJ4qiQQikqrbyIQZ0Z15cUsp5I0feevv0KekiILY9DXKvIvAsE/PzqXO5UtDjQS3QpKpq0fbd+5DeJghcQFGQQrS6gA4D2i/c6ejoITk/nTnDZhHCYDLQ0iYhbGZznDeQXQZiS5uU0jB7+jQUixWgMUwsJQBVNXUNVdU1dVJKsNmQBiMYLUgjCGM7A6IGXzM03KHwJ8EpnpFuZ+H4PMaOHcOAgVbsDU7CeSGERAgRdWQX0k6q2Ux+3khKy44jgxwAakL4I4CammN2T5MTaQSMJrCZwGgBVQuC0M+AzibtRJQVQJ3Xy4d7drN55yGKSqooPlLCmNGjGT7UhsEQ7oYQAqs13XLn7TNvGJt7zeCM7DQ+3FQeVJQCAGm2WPZu+Kik3Ol0kZ5hA7MZjFbADGkaJsMVYECn8/3IxtVlJe/vKuSTfaXUN7aCLRCb0XR506e7Nm1al3H9tTns/WoXfp8vTCZ32DDL3ZNvsnblLj1AkOHg9/uvAZTaOrtbqCporLuAQUszmiCocv1l6QN5JGUHdOeAlTv2H2Tt1sO4vD6MLVehxrT2H5uNtNR0rGnpOOrqcHs8YeKa+vpgkQZDmMqUjIxxg9JT0s0WcxcWXkJKFEXj/KUaXJ4mLBYzqhCgaFp/RU8GiR3QeQtotJGtlMzJH8qSGVkYbWaI9FlVUV0urp04gYHDhlFeWoqjvi5IJFgxq4oSplJVVLS7dhdNGTeGvJF2rGYj8rwnOJ0e9u45QlZWBgNGD2P/toPQxoIgJzQGeIkMdLLhftNnDWPO9dsZPTidl5dNQgx1MnBICrZ0K+4L9RSt30r5oUOkDxrEpHnzGFYwlvqTpyj7/ABNjjrGTb6e1IxMAIx+vxj/47kVYuQI8fXBfZzdfwC/z8vxwmKOl5Ri1eCO28bjdbpgw2Fo8EKzFtEHnJMtAu4JfO8pA+ItiFrkS93tpcbewr0PbeHwkWp++dQk7rtxKFJKak9d4KP1RezbsBtMBu5e+gBT75yNaYCF3Ftu5qtN71G0ej1F6zaRkz+a7JEjuGnOTAZnDrRs2rqL3Xs/57rrJ2IbNgSAtNRUJt86CZqbOLzzS+RHJ+EiTIxo8L8o9HkW6MhlLFLH49ZYu+4QX+4u5eknbuXmCYNba4+Wcl/RTjZs3IPBbOaBR+9nwu2TEAHF1NQ0Zt57D1NmT6dw0xbeXbeZD9e8y/b1GymYcSs/e+5lsrKyuGX6TCbPmo7BaAgqa7GmkTd1EiXFh6jZf6q1DxIB9HkMiHwRCg+GckkTN46283VxJevXFjJz6jCoqqPwgy/ZtmUfBi3QOenZg8i/bXJYp01mM7dOn8b0aVM5ULibta+9zY5NW9i8biNTZ08nb9JN2tRZc7UuOw1KDx3SWECYPjoJkTY+Xpsjx1tDC5yc93j47XO7yB6WxrJFE2G/xLu/mA+2H9J8PnCd1WopXPB/dlYW9z1wL3fcNYstb69j3ctvs/X9j/ho03Y0i0UHXsgwjOij6A+XsaD/KUpHe1CRSgGxCsK3Nj7aUcq+A5X8dtkkbINs8J4Lw7EGPF4vpI8BU1r4sxroqDUlhZlT8hmdm8Mvn1zDn9e8y44ftmC7xnDhHj35aNcxRvDTpwxQ24Hk1CJYf+gsOaMymTZ5OGw9BjYTGK2gSeoaXBjtF/Tl22RFkEPimoypzL9nJvZzVdh86oVvFxfaYu+vGBDLhWBdgdvj4+jJWu6dlQ0Xm+GjY2CzgCUF1GYuH63E5vOBFOHPgRgQxpSUFCZNnoTf7WFg/rmz4aImfp8BLY0nN4hWlde3YLVZmDM2DSoPQNanmr/X1CNbXKR4W9pfb6JbSghVAVVVcbmcVJ06hRCC/rYtFm8MNAJ/SkYBGpAV0Olu9nPBXs+NhiZIGRbYyahvRC0/zbmzFaSmDQxjQOjUtwdBKSUNdbUcLS4mJTWFK/v+FYUuL4O5QJWiKKahQwYr3qYm3C4X41VFNIg0pu9bxntnVgIPgh9wASlGiRh/Dun7Lq7+v48J9lF5OCqPU15aiiEllawRI7BYrec9Xq/XYjLRZX8QQiClxG63U1FejsvpxGiysGHVKrZ//B5Wp/0vwJPfCgMkYHc6nS/+/WN/nfYm5Q9JwZiSSuFb77B5zVrOHC1BU1UyhmThnfQkdp+TlmoHTc9tpeGZNZxav5njux5HGM1UlJygyeGgrqYal9MBoSN7hQhtb/c0ORgxKouKk+WcPlXB/g8/wlxRo+1t3vSv6KXvlwxoAKpcnpb7dh484iwrP0nWkKHYamvZ9sbbrFu5ioqS49hsTJtzO0/8Zz1Dps/AJAzsPfLfVGdfwrDsRp+7AbvhELbxr9JkKGfnts/wejz4/T79ihCMFQGVHq+XjKwsqqvOcqa8nMJduzljP33B4m1aMHnmI1L3SwYA1QAlNfb6n+7Yc6CpprYWl9fP9vXvs+JPz1B2+CgGk4n77p7B4rnTwGLG427GYBIYrRYcBiduk4/aSQdAHcvZwQUMlrvImnaIw15J4UcbaPF4aHI4cTTYcTgcuFxOHPX1nK+s5NTxUuqqznLh7Fky6xuc45rdFOJhT+O8vvk9QJG/p+nS+59+WXx20ZLFt7TU1X2we8dOGhwOfD4ft06dxLw5M8C9F4ypXB6RjS3Vj76LFqvQbBBmG8X7vuS7o/J57/jVeDxeCrdvD21V2xNXFpP3Vx++zQj9oQU1JuWA+MnvN5pVVfX7fT4UVWVyQQF/WP4gDBwITgWaW+CMjw4ZoJ8kEXSOLB3JwtH7OV+7jnEZazDX7abR5yHLfoK9Xh8DZSXIwIRJQLVeVUcw9jwGnLZI1fXjJ194bPHtOY0OB9XnzjFkaBbprVd/DheR1uDKItrPJ0o1E8WrX+I4l8Y/L+dScO4QQ/xGlKSHJfrzgYcBVVMNnz25eO7QOocDo9mMNTX18s7GOQs8bhYgpaS6cQGnLixDlWoXMUH7PwYMvMJLIWEF7kI0WSK6T2bToyFmMCCm0lGKAl8IXnDGdFpfLFiRvYc4Kz+JxD3TQjNUEQppISQwA+K5RB8wwbmPiGdnhPMdMCCB4hHlYw0g9l0gXh/6yp4uQr8i1l+V+oo9F4FYBiYj03sMEG3f++LgJDRPRaElcOOCHWsIIemPLOh79MSA+OMPf/aUCX3LgKixXqbKWAf3nJ49Z0BPrQseP53Grl4kBsS/GnRM9xhqkjMkefSYAXoU3Z84X4B7EQJjHZ9MrxHlEjAr8QxpLxe3zjiXsXj3gJiHJYrw50AXdC6T2MzodNQa7Pis0WsGxNQdzYMYbmOQCVkQjxuRgYuQcD1ASNmINwDR3sQYRmK5TD8loJE4o0/4vdcjADodgBwTSI7bw8wD9LdXLxgQ8aDw6WQkBiTqQTyyodH8WzugT6KA3p3vt0ZEBCBu5kjMgLa0KGZ0EgO6iEOxB9QJFXcVBHTd9yYLOsP/APSqjJQQpDQ/AAAAAElFTkSuQmCC
"""

# Set the appearance mode and default color theme
ctk.set_appearance_mode("Dark")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

class RedirectText:
    """Class to redirect stdout to ScrolledText widget"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = io.StringIO()

    def write(self, string):
        self.buffer.write(string)
        self.text_widget.configure(state=tk.NORMAL)
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state=tk.DISABLED)
        
    def flush(self):
        pass

class InventoryCheckerGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Southern Starz Inventory Checker")
        self.geometry("900x650")
        self.minsize(700, 500)
        
        # Set window icon from base64 data (with better error handling)
        self.set_app_icon()
        
        # Variables
        self.pdf_path = None
        self.output_file = None
        self.processing = False
        self.progress_queue = []  # Queue for progress updates
        
        # Create main container with padding
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create UI elements
        self.create_ui()
    
    def set_app_icon(self):
        """Set the application icon with robust error handling"""
        try:
            # Skip icon setting if PIL is not available
            from PIL import Image
            
            # Use a try/except block for the entire icon setting process
            try:
                # Decode base64 icon
                icon_data = base64.b64decode(APP_ICON_BASE64)
                icon_image = Image.open(BytesIO(icon_data))
                
                # Save as temporary file for Tkinter to use
                temp_icon_path = "temp_icon.ico"
                icon_image.save(temp_icon_path)
                
                # Try to set the icon
                try:
                    self.iconbitmap(temp_icon_path)
                except Exception as e:
                    print(f"Could not set icon bitmap: {e}")
                
                # Clean up the temporary file
                try:
                    os.remove(temp_icon_path)
                except Exception as e:
                    print(f"Could not remove temp icon file: {e}")
            except Exception as e:
                print(f"Error processing icon: {e}")
        except ImportError:
            print("PIL module not available for icon processing")
    
    def create_ui(self):
        """Create all UI elements"""
        # Top section with title and description
        header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Title
        title_label = ctk.CTkLabel(
            header_frame, 
            text="Southern Starz Inventory Checker",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        title_label.pack(anchor="w")
        
        # Description
        desc_label = ctk.CTkLabel(
            header_frame,
            text="Check inventory against website products and generate reports",
            font=ctk.CTkFont(size=14),
            text_color=("gray70", "gray70")
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # File selection section with a nice border
        file_section = ctk.CTkFrame(self.main_container, corner_radius=10)
        file_section.pack(fill=tk.X, pady=(0, 15))
        
        # Add file selection label 
        file_label_container = ctk.CTkFrame(file_section, fg_color="transparent")
        file_label_container.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        file_label = ctk.CTkLabel(
            file_label_container, 
            text="PDF Inventory File",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        file_label.pack(anchor="w")
        
        # Container for file path and browse button
        file_container = ctk.CTkFrame(file_section, fg_color="transparent")
        file_container.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        # Entry for displaying selected file
        self.file_entry = ctk.CTkEntry(
            file_container,
            height=38,
            placeholder_text="Select a PDF inventory file to process"
        )
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Browse button
        browse_btn = ctk.CTkButton(
            file_container, 
            text="Browse",
            height=38,
            font=ctk.CTkFont(weight="bold"),
            command=self.browse_file
        )
        browse_btn.pack(side=tk.RIGHT)
        
        # Action and status section
        action_section = ctk.CTkFrame(self.main_container, corner_radius=10)
        action_section.pack(fill=tk.X, pady=(0, 15))
        
        # Action buttons and progress container
        action_container = ctk.CTkFrame(action_section, fg_color="transparent")
        action_container.pack(fill=tk.X, padx=15, pady=15)
        
        # Process button
        self.process_btn = ctk.CTkButton(
            action_container, 
            text="Process Inventory",
            height=38,
            font=ctk.CTkFont(weight="bold"),
            command=self.process_inventory
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download button (initially disabled)
        self.download_btn = ctk.CTkButton(
            action_container, 
            text="Download Excel",
            height=38,
            font=ctk.CTkFont(weight="bold"),
            state=tk.DISABLED,
            command=self.download_excel
        )
        self.download_btn.pack(side=tk.LEFT)
        
        # Status and progress on the right
        status_container = ctk.CTkFrame(action_container, fg_color="transparent")
        status_container.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Status label
        self.status_label = ctk.CTkLabel(
            status_container, 
            text="Ready",
            font=ctk.CTkFont(size=13),
            text_color=("gray60", "gray60")
        )
        self.status_label.pack(side=tk.RIGHT)
        
        # Progress bar section
        progress_section = ctk.CTkFrame(self.main_container, corner_radius=10)
        progress_section.pack(fill=tk.X, pady=(0, 15))
        
        # Progress container
        progress_container = ctk.CTkFrame(progress_section, fg_color="transparent")
        progress_container.pack(fill=tk.X, padx=15, pady=15)
        
        # Status and percentage label container
        status_percent_container = ctk.CTkFrame(progress_container, fg_color="transparent")
        status_percent_container.pack(fill=tk.X, pady=(0, 5))
        
        # Status message on left
        self.progress_status = ctk.CTkLabel(
            status_percent_container,
            text="Ready",
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        self.progress_status.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Percentage on right
        self.progress_percent = ctk.CTkLabel(
            status_percent_container,
            text="0%",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray60"),
            anchor="e"
        )
        self.progress_percent.pack(side=tk.RIGHT)
        
        # Progress bar below status/percentage
        self.progress_bar = ctk.CTkProgressBar(
            progress_container,
            width=800,
            height=15,
            corner_radius=5
        )
        self.progress_bar.pack(fill=tk.X, expand=True)
        self.progress_bar.set(0)  # Initialize to 0%
        
        # Log section
        log_section = ctk.CTkFrame(self.main_container, corner_radius=10)
        log_section.pack(fill=tk.BOTH, expand=True)
        
        # Log header
        log_header = ctk.CTkFrame(log_section, fg_color="transparent")
        log_header.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        log_title = ctk.CTkLabel(
            log_header, 
            text="Process Log",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        log_title.pack(anchor="w")
        
        # Log content container
        log_container = ctk.CTkFrame(log_section, fg_color="transparent")
        log_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        # Custom-styled log text area
        self.log_text = ctk.CTkTextbox(
            log_container,
            wrap="word",
            height=120,
            corner_radius=6
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)
        
        # Redirect stdout to log_text
        self.text_redirect = RedirectText(self.log_text)
        sys.stdout = self.text_redirect
    
    def update_progress(self, percentage, message=None):
        """Update progress bar and status message"""
        # Ensure percentage is between 0 and 1
        normalized = min(1.0, max(0.0, percentage / 100))
        self.progress_bar.set(normalized)
        self.progress_percent.configure(text=f"{int(percentage)}%")
        
        if message:
            self.progress_status.configure(text=message)
            self.status_label.configure(text=message)  # Update the top status too
            
        # Update UI immediately
        self.update_idletasks()
    
    def browse_file(self):
        """Open file dialog to select PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select PDF Inventory File",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if file_path:
            self.pdf_path = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.status_label.configure(text="PDF file selected")
            
    def process_inventory(self):
        """Process the inventory file in a separate thread"""
        if not self.pdf_path:
            self.status_label.configure(text="No PDF file selected")
            return
        
        if self.processing:
            return
        
        self.processing = True
        self.output_file = None
        self.download_btn.configure(state=tk.DISABLED)
        self.process_btn.configure(state=tk.DISABLED)
        self.status_label.configure(text="Processing...")
        
        # Reset progress
        self.update_progress(0, "Starting...")
        
        # Start processing thread
        threading.Thread(target=self._process_thread, daemon=True).start()
    
    def _process_thread(self):
        """Background thread for processing inventory"""
        try:
            # Clear log
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete("1.0", tk.END)
            self.log_text.configure(state=tk.DISABLED)
            
            # Reset progress bar
            self.update_progress(0, "Starting...")
            
            print(f"Starting inventory check at {datetime.now().strftime('%H:%M:%S')}")
            print(f"Processing PDF: {self.pdf_path}")
            print("-" * 50)
            
            # Extract data from PDF
            self.update_progress(2, "Reading PDF file...")
            checker = InventoryChecker()
            df = checker.extract_pdf_data(self.pdf_path)
            if df is None or df.empty:
                raise Exception("No data extracted from PDF")
            
            print(f"\nFound {len(df)} products in PDF")
            
            # Setup Selenium and prepare for web access - Stage 1 begins (0-15%)
            self.update_progress(5, "Setting up web browser...")
            checker.setup_selenium()
            
            # Get all unique producers
            unique_producers = df['Producer'].unique()
            producers = [p for p in unique_producers if p != "UNKNOWN"]
            products_total = len(df[df['Producer'] != "UNKNOWN"])
            
            # Get all producer web products first (this is the "finding links" stage)
            all_website_products = {}
            for idx, producer in enumerate(producers):
                # Calculate progress for stage 1 (0-15%)
                progress_pct = 5 + int((idx / len(producers)) * 10)  # 5-15% range
                self.update_progress(progress_pct, f"Finding products for {producer}...")
                print(f"\nProcessing {producer}...")
                
                # Get producer products
                website_products = checker.get_producer_products(producer)
                all_website_products[producer] = website_products
                
            # Stage 2: Checking marketing assets (16-85%)
            self.update_progress(15, "Checking marketing assets...")
            
            # Add a progress callback to track progress during the asset checking phase
            def progress_callback(current, total, producer=None):
                if total > 0:
                    # Calculate progress within stage 2 (15-85%, so 70% range)
                    stage_progress = (current / total) * 70
                    # Add the base progress from stage 1 (15%)
                    overall_progress = 15 + stage_progress
                    
                    message = f"Checking assets ({int(overall_progress)}%)"
                    if producer:
                        message = f"Checking {producer} products..."
                    
                    self.update_progress(overall_progress, message)
            
            # Pass the callback to the inventory checker to provide progress updates
            checker.set_progress_callback(progress_callback)
            
            # Run the actual process_inventory method which will check marketing assets
            success = checker.process_inventory(self.pdf_path)
            
            if not success:
                self.update_progress(85, "Processing failed")
                print(f"❌ Processing failed")
                self.status_label.configure(text="Processing failed")
                return
                
            # Stage 3: Creating report (86-100%)
            self.update_progress(85, "Creating Excel report...")
            
            # Find the generated Excel file
            date_string = checker.current_date if hasattr(checker, 'current_date') else datetime.now().strftime('%m%d%y')
            possible_locations = [
                os.path.dirname(os.path.abspath(self.pdf_path)),  # PDF directory
                os.getcwd(),  # Current working directory
                os.path.dirname(os.path.abspath(__file__))  # Script directory
            ]
            
            # Update progress as we format and save the Excel file
            self.update_progress(93, "Formatting Excel report...")
            
            # Look for the Excel file in each possible location
            excel_path = None
            for location in possible_locations:
                base_file = f'inventory_report_{date_string}.xlsx'
                path = os.path.join(location, base_file)
                if os.path.exists(path):
                    excel_path = path
                    self.update_progress(98, "Excel file created")
                    break
            
            # If not found, try numbered versions
            if not excel_path:
                self.update_progress(95, "Searching for generated file...")
                for location in possible_locations:
                    for i in range(1, 11):
                        numbered_file = f'inventory_report_{date_string}_{i}.xlsx'
                        path = os.path.join(location, numbered_file)
                        if os.path.exists(path):
                            excel_path = path
                            self.update_progress(98, "Excel file created")
                            break
                    if excel_path:
                        break
            
            # Final progress update - 100%
            if excel_path:
                self.update_progress(100, "Processing complete!")
                self.output_file = excel_path
                print("-" * 50)
                print(f"✅ Process complete! Excel report generated at:")
                print(f"{excel_path}")
                self.status_label.configure(text="Processing complete!")
                self.download_btn.configure(state=tk.NORMAL)
            else:
                self.update_progress(100, "Excel file not found")
                print("-" * 50)
                print(f"❌ Error: Excel file not found. Searched in:")
                for loc in possible_locations:
                    print(f"  - {loc}")
        
        except Exception as e:
            self.update_progress(100, f"Error: {str(e)}")
            print(f"❌ Error: {str(e)}")
        finally:
            self.processing = False
            self.process_btn.configure(state=tk.NORMAL)
            
    def download_excel(self):
        """Open the generated Excel file"""
        if not self.output_file or not os.path.exists(self.output_file):
            self.status_label.configure(text="No output file available")
            return
        
        try:
            # On Windows, this opens the file with the default application
            os.startfile(self.output_file)
            self.status_label.configure(text="Excel file opened")
        except:
            # If startfile is not available (non-Windows), try another method
            try:
                import subprocess
                # Try platform-specific open commands
                if sys.platform == 'darwin':  # macOS
                    subprocess.call(['open', self.output_file])
                else:  # Linux and others
                    subprocess.call(['xdg-open', self.output_file])
                self.status_label.configure(text="Excel file opened")
            except:
                self.status_label.configure(text="Could not open Excel file")

def main():
    try:
        # Create the application
        app = InventoryCheckerGUI()
        app.mainloop()
    except Exception as e:
        print(f"Error starting application: {str(e)}")
        
        # Display an error message box
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Error", f"Failed to start application: {str(e)}\n\nTrying minimal mode...")
        except:
            # If we can't even show a message box, just print
            print("Could not display error message box")
            print("Trying minimal mode...")
        
        # Try to start in minimal mode (without custom UI components)
        try:
            root = tk.Tk()
            root.title("Southern Starz Inventory Checker (Minimal Mode)")
            root.geometry("800x600")
            
            frame = tk.Frame(root, padx=20, pady=20)
            frame.pack(fill=tk.BOTH, expand=True)
            
            label = tk.Label(frame, text="Southern Starz Inventory Checker", font=("Arial", 16))
            label.pack(pady=(0, 10))
            
            file_frame = tk.Frame(frame)
            file_frame.pack(fill=tk.X, pady=10)
            
            tk.Label(file_frame, text="PDF File:").pack(side=tk.LEFT, padx=(0, 10))
            entry = tk.Entry(file_frame)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
            
            pdf_path = [None]  # Use a list to allow modification in nested function
            
            def browse():
                file = filedialog.askopenfilename(
                    title="Select PDF Inventory File",
                    filetypes=[("PDF Files", "*.pdf")]
                )
                if file:
                    entry.delete(0, tk.END)
                    entry.insert(0, file)
                    pdf_path[0] = file
                    
            def process():
                if not pdf_path[0]:
                    messagebox.showinfo("Error", "Please select a PDF file first")
                    return
                
                process_btn.config(state=tk.DISABLED)
                status_label.config(text="Processing... This may take a while.")
                
                # Run in separate thread to avoid freezing UI
                def run_process():
                    try:
                        checker = InventoryChecker()
                        checker.process_inventory(pdf_path[0])
                        root.after(0, lambda: status_label.config(text="Complete! Check for Excel file in the same folder."))
                        root.after(0, lambda: process_btn.config(state=tk.NORMAL))
                    except Exception as e:
                        root.after(0, lambda: status_label.config(text=f"Error: {str(e)}"))
                        root.after(0, lambda: process_btn.config(state=tk.NORMAL))
                
                threading.Thread(target=run_process, daemon=True).start()
                
            browse_btn = tk.Button(file_frame, text="Browse", command=browse)
            browse_btn.pack(side=tk.RIGHT)
            
            process_btn = tk.Button(frame, text="Process Inventory", command=process)
            process_btn.pack(pady=10)
            
            status_label = tk.Label(frame, text="Ready", fg="gray")
            status_label.pack(pady=5)
            
            # Add log text area
            log_frame = tk.Frame(frame)
            log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            log_label = tk.Label(log_frame, text="Process Log")
            log_label.pack(anchor=tk.W)
            
            log_text = ScrolledText(log_frame, height=15)
            log_text.pack(fill=tk.BOTH, expand=True)
            log_text.config(state=tk.DISABLED)
            
            # Redirect stdout
            sys.stdout = RedirectText(log_text)
            
            root.mainloop()
        except Exception as e:
            print(f"Fatal error in minimal mode: {str(e)}")
            
        # Check if PIL is correctly installed (just for diagnostics)
        try:
            from PIL import Image
            print("PIL seems to be installed correctly")
        except ImportError:
            print("PIL (Pillow) is not installed properly. Try running: pip install Pillow")

if __name__ == "__main__":
    main() 