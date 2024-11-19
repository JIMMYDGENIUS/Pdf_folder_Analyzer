import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pypdf import PdfReader
import csv
from datetime import datetime
import pathlib
from docx import Document
import pandas as pd
import logging
from datetime import datetime
import win32com.client  # For better Word page counting
import xlrd  # For Excel metadata
from tkinter import font as tkfont

class PDFAnalyzer:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF Page Size Analyzer")
        self.window.geometry("1024x768")
        
        # Configure window styling
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure colors and styles
        self.window.configure(bg='#f0f0f0')
        self.style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        self.style.configure('Custom.TButton', padding=5)
        
        # Store analysis results
        self.results = []
        
        self.create_gui()
        
    def create_gui(self):
        # Status bar at the top
        self.status_var = tk.StringVar()
        status_label = ttk.Label(self.window, textvariable=self.status_var)
        status_label.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

        # Button frame
        button_frame = ttk.Frame(self.window)
        button_frame.grid(row=1, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

        # Select button
        select_btn = ttk.Button(button_frame, text="Select Folder", 
                              command=self.select_folder, 
                              style='Custom.TButton')
        select_btn.pack(side='left', padx=5)

        # Export button
        export_btn = ttk.Button(button_frame, text="Export to CSV",
                              command=self.export_to_csv,
                              style='Custom.TButton')
        export_btn.pack(side='left', padx=5)

        # Single Summary Frame at the top
        summary_frame = ttk.LabelFrame(self.window, text="Document Analysis Summary")
        summary_frame.grid(row=2, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

        # Create two columns in summary frame
        left_summary = ttk.Frame(summary_frame)
        left_summary.pack(side='left', padx=10, pady=5)
        
        right_summary = ttk.Frame(summary_frame)
        right_summary.pack(side='right', padx=10, pady=5)

        # Summary labels with more detailed information
        self.summary_labels = {
            # Left column - File counts
            'pdf': ttk.Label(left_summary, text="PDF Files: 0"),
            'pdf_pages': ttk.Label(left_summary, text="PDF Pages: 0"),
            'word': ttk.Label(left_summary, text="Word Files: 0"),
            'word_pages': ttk.Label(left_summary, text="Word Pages: 0"),
            'excel': ttk.Label(left_summary, text="Excel Files: 0"),
            'excel_sheets': ttk.Label(left_summary, text="Excel Sheets: 0"),
            
            # Right column - Totals
            'total_files': ttk.Label(right_summary, text="Total Files: 0"),
            'total_pages': ttk.Label(right_summary, text="Total Pages/Sheets: 0")
        }

        # Pack left column labels
        for label in ['pdf', 'pdf_pages', 'word', 'word_pages', 'excel', 'excel_sheets']:
            self.summary_labels[label].pack(anchor='w')

        # Pack right column labels
        for label in ['total_files', 'total_pages']:
            self.summary_labels[label].pack(anchor='e')

        # Folder Breakdown section
        folder_frame = ttk.LabelFrame(self.window, text="Folder Breakdown")
        folder_frame.grid(row=5, column=0, columnspan=2, sticky='nsew', padx=10, pady=5)

        # Add scrollbar for folder breakdown
        folder_scroll = ttk.Scrollbar(folder_frame)
        folder_scroll.pack(side='right', fill='y')

        # Create text widget with scrollbar
        self.folder_text = tk.Text(folder_frame, height=6, wrap=tk.WORD, 
                                 yscrollcommand=folder_scroll.set)
        self.folder_text.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        folder_scroll.config(command=self.folder_text.yview)

        # Configure grid weights to allow proper expansion
        self.window.grid_rowconfigure(5, weight=1)
        self.window.grid_columnconfigure(0, weight=1)

        # Move progress frame to row 3
        progress_frame = ttk.Frame(self.window)
        progress_frame.grid(row=3, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            orient='horizontal', 
            length=300, 
            mode='determinate'
        )
        self.progress_bar.pack(side='left', fill='x', expand=True)

        # Progress label
        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.pack(side='right', padx=5)

        # Treeview
        columns = ('File Name', 'File Path', 'File Type', 'Page', 
                  'Width (mm)', 'Height (mm)', 'Size & Recommended Paper')
        
        column_widths = {
            'File Name': 200,
            'File Path': 300,
            'File Type': 100,
            'Page': 80,
            'Width (mm)': 120,
            'Height (mm)': 120,
            'Size & Recommended Paper': 300
        }
        
        self.tree = ttk.Treeview(self.window, columns=columns, show='headings')
        self.tree.grid(row=4, column=0, sticky='nsew', padx=10, pady=5)

        # Configure treeview styling
        self.style.configure('Custom.Treeview', rowheight=25)
        self.style.configure('Custom.Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        
        # Set column headings and widths
        for col, width in column_widths.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width)
        
        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(self.window, orient=tk.VERTICAL, command=self.tree.yview)
        x_scrollbar = ttk.Scrollbar(self.window, orient=tk.HORIZONTAL, command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # Grid layout for table and scrollbars
        y_scrollbar.grid(row=4, column=1, sticky='ns')
        x_scrollbar.grid(row=5, column=0, sticky='ew')
    
    def select_folder(self):
        folder_path = filedialog.askdirectory(title="Select Folder Containing Documents")
        if folder_path:
            self.status_var.set(f"Analyzing folder: {folder_path}")
            self.window.update()
            self.analyze_folder(folder_path)
            self.status_var.set("Analysis complete")
    
    def analyze_folder(self, folder_path):
        try:
            # Clear previous results
            self.results = []
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Reset progress bar
            self.progress_bar['value'] = 0
            self.progress_label['text'] = "0%"
            
            # First count total files for progress calculation
            total_files = 0
            supported_extensions = {'.pdf', '.docx', '.doc', '.xlsx', '.xls'}
            
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if pathlib.Path(file).suffix.lower() in supported_extensions:
                        total_files += 1
            
            if total_files == 0:
                self.status_var.set("No supported files found")
                return
                
            # Now process the files
            processed_files = 0
            
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = pathlib.Path(root) / file
                    if file_path.suffix.lower() in supported_extensions:
                        self.analyze_file(file_path)
                        processed_files += 1
                        
                        # Update progress
                        progress = (processed_files / total_files) * 100
                        self.progress_bar['value'] = progress
                        self.progress_label['text'] = f"{progress:.1f}%"
                        self.window.update_idletasks()  # Update GUI
                        
            # After processing all files
            self.update_summaries()  # Update both summaries
            
            # Re-enable text widget for updating
            self.folder_text.configure(state='normal')
            self.update_summaries()
            # Make text widget read-only again
            self.folder_text.configure(state='disabled')
            
            self.status_var.set(f"Analysis complete. Found {len(self.results)} results")
            
        except Exception as e:
            print(f"Error in analyze_folder: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Error analyzing folder: {str(e)}")
        finally:
            # Ensure progress bar shows complete
            self.progress_bar['value'] = 100
            self.progress_label['text'] = "100%"
            self.window.update_idletasks()
    
    def analyze_file(self, file_path):
        file_type = file_path.suffix.lower()
        try:
            print(f"Analyzing file: {file_path}")
            if file_type == '.pdf':
                self.analyze_pdf(file_path)
            elif file_type in {'.docx', '.doc'}:
                self.analyze_word(file_path)
            elif file_type in {'.xlsx', '.xls'}:
                self.analyze_excel(file_path)
        except Exception as e:
            print(f"Error analyzing {file_path}: {str(e)}")
            self.add_result(
                file_path=file_path,
                file_type=file_type,
                page='Error',
                width='N/A',
                height='N/A',
                size=f'Error: {str(e)}'
            )
    
    def analyze_pdf(self, file_path):
        """Enhanced PDF analysis with multiple verification methods"""
        try:
            # Primary method using pypdf
            reader = PdfReader(str(file_path))
            
            # Verify PDF is readable
            if reader.is_encrypted:
                try:
                    # Try to read without password (some PDFs are encrypted but don't need password)
                    reader.decrypt('')
                except:
                    logging.error(f"Encrypted PDF cannot be read: {file_path}")
                    raise ValueError("PDF is encrypted and requires password")

            # Multiple methods to verify page count
            reported_pages = len(reader.pages)
            verified_pages = 0
            page_sizes = {}
            
            # Method 1: Try to access each page
            for page_num in range(reported_pages):
                try:
                    page = reader.pages[page_num]
                    if page is not None:
                        verified_pages += 1
                        
                        # Get page dimensions
                        width_mm = round(float(page.mediabox.width) * 0.352778, 1)
                        height_mm = round(float(page.mediabox.height) * 0.352778, 1)
                        
                        # Store unique page sizes
                        size_key = f"{width_mm}x{height_mm}"
                        if size_key not in page_sizes:
                            page_sizes[size_key] = {
                                'width': width_mm,
                                'height': height_mm,
                                'count': 1,
                                'pages': [page_num + 1]
                            }
                        else:
                            page_sizes[size_key]['count'] += 1
                            page_sizes[size_key]['pages'].append(page_num + 1)
                            
                except Exception as e:
                    logging.warning(f"Could not access page {page_num + 1} in {file_path}: {str(e)}")

            # Method 2: Verify using PDF structure (if page count seems wrong)
            if verified_pages < reported_pages:
                try:
                    with open(str(file_path), 'rb') as file:
                        content = file.read()
                        # Look for PDF page markers
                        page_markers = content.count(b'/Type /Page')
                        obj_markers = content.count(b'/Page\n') + content.count(b'/Page\r\n')
                        
                        # Use the most reliable count
                        structure_pages = max(page_markers, obj_markers)
                        if structure_pages > verified_pages:
                            verified_pages = structure_pages
                except Exception as e:
                    logging.error(f"Error in structural analysis of {file_path}: {str(e)}")

            # If we found no pages but the file exists, something is wrong
            if verified_pages == 0 and os.path.getsize(file_path) > 0:
                raise ValueError("Could not verify any pages in apparently non-empty PDF")

            # Determine the most common page size
            if page_sizes:
                main_size = max(page_sizes.values(), key=lambda x: x['count'])
                width_mm = main_size['width']
                height_mm = main_size['height']
                size_name = self.determine_page_size(width_mm, height_mm)
                
                # If there are multiple page sizes, add this to the size name
                if len(page_sizes) > 1:
                    size_ranges = []
                    for size_data in page_sizes.values():
                        page_range = self.format_page_ranges(size_data['pages'])
                        size_mm = f"{size_data['width']}x{size_data['height']}mm"
                        size_ranges.append(f"{size_mm} (pages {page_range})")
                    size_name += f" | Mixed sizes: {'; '.join(size_ranges)}"
            else:
                width_mm = height_mm = 'N/A'
                size_name = 'Unknown'

            # Add the result
            self.add_result(
                file_path=file_path,
                file_type='.pdf',
                page=verified_pages,
                width=width_mm,
                height=height_mm,
                size=size_name
            )

            # Log analysis results
            logging.info(f"PDF Analysis - File: {file_path}, "
                        f"Pages: {verified_pages}, "
                        f"Sizes: {len(page_sizes)}")

        except Exception as e:
            logging.error(f"Error analyzing PDF {file_path}: {str(e)}")
            self.add_result(
                file_path=file_path,
                file_type='.pdf',
                page='Error',
                width='N/A',
                height='N/A',
                size=f'Error: {str(e)}'
            )

    def format_page_ranges(self, page_numbers):
        """Convert a list of page numbers to a condensed range string"""
        if not page_numbers:
            return ""
        
        ranges = []
        start = page_numbers[0]
        prev = start
        
        for num in page_numbers[1:] + [None]:
            if num != prev + 1:
                if start == prev:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{prev}")
                start = num
            prev = num
            
        return ", ".join(ranges)
    
    def determine_page_size(self, width_mm, height_mm):
        # Always use larger dimension as width
        width_mm, height_mm = max(width_mm, height_mm), min(width_mm, height_mm)
        
        # Standard page sizes (width × height in mm)
        sizes = {
            'A0': (841, 1189),
            'A1': (594, 841),
            'A2': (420, 594),
            'A3': (297, 420),
            'A4': (210, 297),
            'A5': (148, 210),
            'Letter': (215.9, 279.4),
            'Legal': (215.9, 355.6)
        }
        
        actual_size = None
        recommended_size = None
        
        # Find actual size with tolerance
        tolerance = 5
        for size_name, (std_width, std_height) in sizes.items():
            if (abs(width_mm - std_width) <= tolerance and 
                abs(height_mm - std_height) <= tolerance):
                actual_size = size_name
                break
        
        # Find recommended size (smallest size that fits the content)
        for size_name, (std_width, std_height) in sorted(sizes.items(), key=lambda x: x[1][0] * x[1][1]):
            if width_mm <= std_width + tolerance and height_mm <= std_height + tolerance:
                recommended_size = size_name
                break
        
        actual_size = actual_size or f"Custom ({width_mm:.1f}×{height_mm:.1f}mm)"
        recommended_size = recommended_size or "Custom (Too large for standard sizes)"
        
        return f"{actual_size} (Print on: {recommended_size})"
    
    def update_summaries(self):
        """Update summary with detailed counts"""
        counts = {
            'pdf_files': 0,
            'pdf_pages': 0,
            'word_files': 0,
            'word_pages': 0,
            'excel_files': 0,
            'excel_sheets': 0,
            'folders': {}
        }

        # Process each result
        for result in self.results:
            file_type = result['file_type'].lower()
            folder_path = str(pathlib.Path(result['file_path']).parent)
            
            # Initialize folder in counts if not exists
            if folder_path not in counts['folders']:
                counts['folders'][folder_path] = {
                    'pdf_files': 0, 'pdf_pages': 0,
                    'word_files': 0, 'word_pages': 0,
                    'excel_files': 0, 'excel_sheets': 0
                }

            # Get page count - now stored directly as number
            page_count = result['page'] if isinstance(result['page'], (int, float)) else 0

            # Update counts based on file type
            if file_type == '.pdf':
                counts['pdf_files'] += 1
                counts['pdf_pages'] += page_count
                counts['folders'][folder_path]['pdf_files'] += 1
                counts['folders'][folder_path]['pdf_pages'] += page_count
            elif file_type in ['.doc', '.docx']:
                counts['word_files'] += 1
                counts['word_pages'] += page_count
                counts['folders'][folder_path]['word_files'] += 1
                counts['folders'][folder_path]['word_pages'] += page_count
            elif file_type in ['.xls', '.xlsx']:
                counts['excel_files'] += 1
                counts['excel_sheets'] += page_count
                counts['folders'][folder_path]['excel_files'] += 1
                counts['folders'][folder_path]['excel_sheets'] += page_count

        # Update summary labels
        self.summary_labels['pdf'].configure(text=f"PDF Files: {counts['pdf_files']}")
        self.summary_labels['pdf_pages'].configure(text=f"PDF Pages: {counts['pdf_pages']}")
        self.summary_labels['word'].configure(text=f"Word Files: {counts['word_files']}")
        self.summary_labels['word_pages'].configure(text=f"Word Pages: {counts['word_pages']}")
        self.summary_labels['excel'].configure(text=f"Excel Files: {counts['excel_files']}")
        self.summary_labels['excel_sheets'].configure(text=f"Excel Sheets: {counts['excel_sheets']}")

        total_files = counts['pdf_files'] + counts['word_files'] + counts['excel_files']
        total_pages = counts['pdf_pages'] + counts['word_pages'] + counts['excel_sheets']
        
        self.summary_labels['total_files'].configure(text=f"Total Files: {total_files}")
        self.summary_labels['total_pages'].configure(text=f"Total Pages/Sheets: {total_pages}")

        # Update folder breakdown text
        self.folder_text.delete('1.0', tk.END)
        for folder, folder_counts in counts['folders'].items():
            folder_total_files = (folder_counts['pdf_files'] + 
                                folder_counts['word_files'] + 
                                folder_counts['excel_files'])
            folder_total_pages = (folder_counts['pdf_pages'] + 
                                folder_counts['word_pages'] + 
                                folder_counts['excel_sheets'])
            
            folder_summary = (
                f"\nFolder: {folder}\n"
                f"  PDF:   {folder_counts['pdf_files']} files, {folder_counts['pdf_pages']} pages\n"
                f"  Word:  {folder_counts['word_files']} files, {folder_counts['word_pages']} pages\n"
                f"  Excel: {folder_counts['excel_files']} files, {folder_counts['excel_sheets']} sheets\n"
                f"  Totals: {folder_total_files} files, {folder_total_pages} pages/sheets\n"
                f"{'-' * 80}\n"
            )
            self.folder_text.insert(tk.END, folder_summary)

    def export_results(self):
        """Enhanced export functionality"""
        if not self.results:
            messagebox.showwarning("Warning", "No results to export!")
            return
            
        # Ask user what to export
        export_options = tk.Toplevel(self.window)
        export_options.title("Export Options")
        
        fields = {
            'file_name': tk.BooleanVar(value=True),
            'file_path': tk.BooleanVar(value=True),
            'file_type': tk.BooleanVar(value=True),
            'page_count': tk.BooleanVar(value=True),
            'dimensions': tk.BooleanVar(value=True),
            'metadata': tk.BooleanVar(value=False)
        }
        
        for field, var in fields.items():
            ttk.Checkbutton(
                export_options, 
                text=field.replace('_', ' ').title(),
                variable=var
            ).pack(anchor='w', padx=5, pady=2)
        
        # Export format selection
        format_var = tk.StringVar(value='csv')
        ttk.Radiobutton(
            export_options, 
            text="CSV", 
            variable=format_var, 
            value='csv'
        ).pack(anchor='w', padx=5)
        ttk.Radiobutton(
            export_options, 
            text="Excel", 
            variable=format_var, 
            value='xlsx'
        ).pack(anchor='w', padx=5)
        
        def do_export():
            selected_fields = [f for f, v in fields.items() if v.get()]
            self.export_to_file(selected_fields, format_var.get())
            export_options.destroy()
        
        ttk.Button(
            export_options, 
            text="Export", 
            command=do_export
        ).pack(pady=10)

    def export_to_file(self, fields, format_type):
        """Export results to selected format"""
        try:
            # Prepare data based on selected fields
            export_data = []
            for result in self.results:
                row = {}
                for field in fields:
                    if field in result:
                        row[field] = result[field]
                    elif field == 'metadata' and 'metadata' in result:
                        row.update(result['metadata'])
                export_data.append(row)
            
            # Get save location
            file_types = [('CSV files', '*.csv')] if format_type == 'csv' else [('Excel files', '*.xlsx')]
            file_path = filedialog.asksaveasfilename(
                defaultextension=f'.{format_type}',
                filetypes=file_types,
                title="Export Results"
            )
            
            if file_path:
                if format_type == 'csv':
                    pd.DataFrame(export_data).to_csv(file_path, index=False)
                else:
                    pd.DataFrame(export_data).to_excel(file_path, index=False)
                    
                messagebox.showinfo("Success", "Results exported successfully!")
                
        except Exception as e:
            logging.error(f"Error exporting results: {str(e)}")
            messagebox.showerror("Error", f"Error exporting results: {str(e)}")

    def show_error_log(self):
        """Display error log window"""
        if not self.errors:
            messagebox.showinfo("Info", "No errors to display")
            return
            
        error_window = tk.Toplevel(self.window)
        error_window.title("Error Log")
        
        # Create text widget with scrollbar
        error_text = tk.Text(error_window, wrap=tk.WORD, width=60, height=20)
        scrollbar = ttk.Scrollbar(error_window, command=error_text.yview)
        error_text.configure(yscrollcommand=scrollbar.set)
        
        error_text.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Insert errors
        for error in self.errors:
            error_text.insert(tk.END, f"File: {error['file']}\n")
            error_text.insert(tk.END, f"Error: {error['error']}\n")
            error_text.insert(tk.END, f"Type: {error['type']}\n")
            error_text.insert(tk.END, "-" * 50 + "\n")
        
        error_text.configure(state='disabled')
    
    def on_closing(self):
        """Handle window closing event"""
        self.window.quit()
        self.window.destroy()
    
    def run(self):
        self.window.mainloop()
    
    def reset_progress(self):
        """Reset progress bar and label"""
        self.progress_bar['value'] = 0
        self.progress_label['text'] = "0%"
        self.window.update_idletasks()
    
    def update_progress(self, current, total):
        """Update progress bar and label"""
        progress = (current / total) * 100 if total > 0 else 0
        self.progress_bar['value'] = progress
        self.progress_label['text'] = f"{progress:.1f}%"
        self.window.update_idletasks()
    
    def export_to_csv(self):
        """Export analysis results to CSV file"""
        if not self.results:
            messagebox.showwarning("Warning", "No results to export!")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            title="Export Results"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=[
                        'file_name', 'file_path', 'file_type', 'page', 
                        'width', 'height', 'size'
                    ])
                    writer.writeheader()
                    writer.writerows(self.results)
                messagebox.showinfo("Success", "Results exported successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Error exporting results: {str(e)}")
    
    def add_result(self, file_path, file_type, page, width, height, size):
        """Add a result to both the results list and treeview"""
        try:
            # Get the file name from the path
            if isinstance(file_path, pathlib.Path):
                file_name = file_path.name
                file_path_str = str(file_path)
            else:
                file_name = pathlib.Path(file_path).name
                file_path_str = str(file_path)

            # Create result dictionary
            result = {
                'file_name': file_name,
                'file_path': file_path_str,
                'file_type': file_type,
                'page': page,
                'width': width,
                'height': height,
                'size': size
            }
            
            # Add to results list
            self.results.append(result)
            
            # Add to treeview
            self.tree.insert('', tk.END, values=(
                file_name,
                file_path_str,
                file_type,
                page,
                width if width != 'N/A' else 'N/A',
                height if height != 'N/A' else 'N/A',
                size
            ))
            
            # Update the GUI
            self.window.update_idletasks()
            
        except Exception as e:
            print(f"Error in add_result: {str(e)}")
            raise
    
    def analyze_word(self, file_path):
        """Improved Word document analysis"""
        try:
            # Try using COM object for accurate page count
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(file_path))
            page_count = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
            
            # Get document properties
            width_mm = round(doc.PageSetup.PageWidth * 0.352778, 1)
            height_mm = round(doc.PageSetup.PageHeight * 0.352778, 1)
            
            doc.Close()
            word.Quit()
            
            result_size = self.determine_page_size(width_mm, height_mm)
            
            self.add_result(
                file_path=file_path,
                file_type='.docx',
                page=page_count,
                width=width_mm,
                height=height_mm,
                size=result_size,
                metadata={'actual_page_count': True}
            )
            
        except Exception as e:
            logging.error(f"Error analyzing Word document {file_path}: {str(e)}")
            # Fallback to python-docx
            self.analyze_word_fallback(file_path)

    def analyze_excel(self, file_path):
        """Enhanced Excel analysis"""
        try:
            xl = pd.ExcelFile(file_path)
            metadata = {
                'sheet_count': len(xl.sheet_names),
                'sheets': {}
            }
            
            # Analyze each sheet
            for sheet_name in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name)
                metadata['sheets'][sheet_name] = {
                    'rows': len(df),
                    'columns': len(df.columns),
                    'non_empty_cells': df.count().sum()
                }
            
            self.add_result(
                file_path=file_path,
                file_type='.xlsx',
                page=len(xl.sheet_names),
                width='N/A',
                height='N/A',
                size='Excel Workbook',
                metadata=metadata
            )
            
        except Exception as e:
            logging.error(f"Error analyzing Excel file {file_path}: {str(e)}")
            self.errors.append({
                'file': file_path,
                'error': str(e),
                'type': 'excel'
            })

if __name__ == "__main__":
    app = PDFAnalyzer()
    app.run() 