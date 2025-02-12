import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
import os
import threading
import queue
from pathlib import Path
import sys

from MCQQuestionSplitter import MCQQuestionSplitter

class LogRedirector:
    def __init__(self, text_widget, queue):
        self.text_widget = text_widget
        self.queue = queue

    def write(self, str):
        self.queue.put(str)

    def flush(self):
        pass

class MCQSplitterGUI:
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("MCQ PDF to PowerPoint Converter")
        self.window.geometry("900x570")
        
        # Set theme
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        # Queue for log messages
        self.log_queue = queue.Queue()
        self.setup_gui()
        self.check_log_queue()

    def setup_gui(self):
        # Create main frame with scrollable content
        self.main_frame = ctk.CTkScrollableFrame(self.window)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Title
        title_label = ctk.CTkLabel(
            self.main_frame, 
            text="MCQ PDF to PowerPoint Converter",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=10)

        # Mode selection
        self.setup_mode_selection()
        
        # Single file section
        self.single_file_frame = ctk.CTkFrame(self.main_frame)
        self.single_file_frame.pack(fill=tk.X, padx=5, pady=5)
        self.setup_single_file_section()
        
        # Batch processing section
        self.batch_frame = ctk.CTkFrame(self.main_frame)
        self.batch_frame.pack(fill=tk.X, padx=5, pady=5)
        self.setup_batch_processing()
        
        # Time settings section
        self.setup_time_settings()
            
        # Log section
        self.setup_log_section()
        
        # Process button
        self.process_button = ctk.CTkButton(
            self.main_frame,
            text="Start Processing",
            command=self.process_files,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.process_button.pack(pady=10)

        # Initially hide batch frame
        self.batch_frame.pack_forget()

        # Redirect stdout to our log widget
        sys.stdout = LogRedirector(self.log_text, self.log_queue)

    def setup_mode_selection(self):
        self.mode_frame = ctk.CTkFrame(self.main_frame)
        self.mode_frame.pack(fill=tk.X, padx=5, pady=5)
        
        mode_label = ctk.CTkLabel(self.mode_frame, text="Processing Mode:", font=ctk.CTkFont(weight="bold"))
        mode_label.pack(side=tk.LEFT, padx=5)
        
        self.mode_var = ctk.StringVar(value="single")
        
        single_radio = ctk.CTkRadioButton(
            self.mode_frame,
            text="Single File",
            variable=self.mode_var,
            value="single",
            command=self.toggle_mode
        )
        single_radio.pack(side=tk.LEFT, padx=20)
        
        batch_radio = ctk.CTkRadioButton(
            self.mode_frame,
            text="Batch Processing",
            variable=self.mode_var,
            value="batch",
            command=self.toggle_mode
        )
        batch_radio.pack(side=tk.LEFT, padx=20)

    def setup_single_file_section(self):
        # Input file
        self.create_file_selection(
            self.single_file_frame,
            "Input PDF File:",
            "input_path",
            self.browse_input,
            "Select a PDF file containing MCQ questions"
        )
        
        # Output file
        self.create_file_selection(
            self.single_file_frame,
            "Output PPTX File:",
            "output_path",
            self.browse_output,
            "Select where to save the PowerPoint presentation"
        )

    def setup_batch_processing(self):
        # Input directory
        self.create_file_selection(
            self.batch_frame,
            "Input PDF Directory:",
            "batch_input_path",
            self.browse_batch_input,
            "Select directory containing PDF files"
        )
        
        # Output directory
        self.create_file_selection(
            self.batch_frame,
            "Output Directory:",
            "batch_output_path",
            self.browse_batch_output,
            "Select directory for output PPTX files"
        )

    def create_file_selection(self, parent, label_text, path_attr, browse_command, help_text):
        frame = ctk.CTkFrame(parent)
        frame.pack(fill=tk.X, padx=5, pady=5)
        
        label = ctk.CTkLabel(frame, text=label_text, width=120)
        label.pack(side=tk.LEFT, padx=5)
        
        entry = ctk.CTkEntry(frame)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        setattr(self, path_attr, entry)
        
        browse_btn = ctk.CTkButton(
            frame,
            text="Browse",
            command=browse_command,
            width=100
        )
        browse_btn.pack(side=tk.LEFT, padx=5)
        
        help_btn = ctk.CTkButton(
            frame,
            text="?",
            width=30,
            command=lambda: self.show_help(help_text)
        )
        help_btn.pack(side=tk.LEFT, padx=5)

    def setup_time_settings(self):
        time_frame = ctk.CTkFrame(self.main_frame)
        time_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Add checkbox for enabling/disabling timing
        self.timing_enabled = ctk.BooleanVar(value=True)
        timing_checkbox = ctk.CTkCheckBox(
            time_frame,
            text="Enable automatic slide timing",
            variable=self.timing_enabled,
            command=self.toggle_timing_entry
        )
        timing_checkbox.pack(side=tk.LEFT, padx=5)
        
        # Time entry section
        self.time_entry_frame = ctk.CTkFrame(time_frame)
        self.time_entry_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        time_label = ctk.CTkLabel(self.time_entry_frame, text="Time per slide:")
        time_label.pack(side=tk.LEFT, padx=5)
        
        self.time_entry = ctk.CTkEntry(self.time_entry_frame, width=100)
        self.time_entry.insert(0, "15")
        self.time_entry.pack(side=tk.LEFT, padx=5)
        
        seconds_label = ctk.CTkLabel(self.time_entry_frame, text="seconds")
        seconds_label.pack(side=tk.LEFT, padx=5)
        
        help_button = ctk.CTkButton(
            time_frame,
            text="?",
            width=30,
            command=lambda: self.show_help("Set how long each question will be displayed. Disable to control slides manually.")
        )
        help_button.pack(side=tk.LEFT, padx=5)

    def setup_log_section(self):
        log_frame = ctk.CTkFrame(self.main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        log_label = ctk.CTkLabel(log_frame, text="Processing Log:", font=ctk.CTkFont(weight="bold"))
        log_label.pack(anchor=tk.W, padx=5, pady=2)
        
        self.log_text = ctk.CTkTextbox(log_frame, height=200)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def toggle_mode(self):
        if self.mode_var.get() == "batch":
            self.single_file_frame.pack_forget()
            self.batch_frame.pack(after=self.mode_frame, fill=tk.X, padx=5, pady=5)
        else:
            self.batch_frame.pack_forget()
            self.single_file_frame.pack(after=self.mode_frame, fill=tk.X, padx=5, pady=5)

    def toggle_timing_entry(self):
        """Enable/disable time entry based on checkbox state"""
        if self.timing_enabled.get():
            self.time_entry.configure(state="normal")
        else:
            self.time_entry.configure(state="disabled")

    def show_help(self, message):
        messagebox.showinfo("Help", message)

    def browse_input(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            self.input_path.delete(0, tk.END)
            self.input_path.insert(0, file_path)
            
            # Auto-fill output path if empty
            if not self.output_path.get():
                output_path = Path(file_path).with_suffix('.pptx')
                self.output_path.delete(0, tk.END)
                self.output_path.insert(0, str(output_path))

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if file_path:
            self.output_path.delete(0, tk.END)
            self.output_path.insert(0, file_path)

    def browse_batch_input(self):
        dir_path = filedialog.askdirectory(title="Select Input Directory")
        if dir_path:
            self.batch_input_path.delete(0, tk.END)
            self.batch_input_path.insert(0, dir_path)

    def browse_batch_output(self):
        dir_path = filedialog.askdirectory(title="Select Output Directory")
        if dir_path:
            self.batch_output_path.delete(0, tk.END)
            self.batch_output_path.insert(0, dir_path)

    def process_single_file(self, pdf_path, output_path, seconds):
        try:
            self.log_text.insert(tk.END, f"\nProcessing {os.path.basename(pdf_path)}...\n")
            
            # Pass None for seconds if timing is disabled
            actual_seconds = seconds if self.timing_enabled.get() else None
            converter = MCQQuestionSplitter(slide_duration=actual_seconds)
            
            converter.convert_pdf_to_slides(pdf_path, output_path)
            self.log_text.insert(tk.END, f"Successfully processed {pdf_path}\n")
            return True
        except Exception as e:
            self.log_text.insert(tk.END, f"Error processing {pdf_path}: {str(e)}\n")
            return False

    def process_files(self):
        # Only validate seconds if timing is enabled
        if self.timing_enabled.get():
            try:
                seconds = int(self.time_entry.get())
                if seconds <= 0:
                    raise ValueError("Seconds must be positive")
            except ValueError as e:
                messagebox.showerror("Error", "Please enter a valid number of seconds")
                return
        else:
            seconds = None

        self.process_button.configure(state="disabled")

        if self.mode_var.get() == "batch":
            input_dir = self.batch_input_path.get()
            output_dir = self.batch_output_path.get()
            
            if not input_dir or not output_dir:
                messagebox.showerror("Error", "Please select input and output directories")
                self.process_button.configure(state="normal")
                return
            
            def process_batch():
                pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
                total_files = len(pdf_files)
                
                for i, file in enumerate(pdf_files, 1):
                    pdf_path = os.path.join(input_dir, file)
                    output_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}_mcq.pptx")
                    
                    self.process_single_file(pdf_path, output_path, seconds)
                
                self.log_text.insert(tk.END, "\nBatch processing completed\n")
                self.process_button.configure(state="normal")
                
                messagebox.showinfo("File processing completed successfully!")

            threading.Thread(target=process_batch, daemon=True).start()
        
        else:
            input_path = self.input_path.get()
            output_path = self.output_path.get()
            
            if not input_path or not output_path:
                messagebox.showerror("Error", "Please select input and output files")
                self.process_button.configure(state="normal")
                return
            
            def process_single():
                success = self.process_single_file(input_path, output_path, seconds)
                self.process_button.configure(state="normal")
                
                if success:
                    messagebox.showinfo("Success", "File processing completed successfully!")
            
            threading.Thread(target=process_single, daemon=True).start()

    def check_log_queue(self):
        while True:
            try:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, message)
                self.log_text.see(tk.END)
            except queue.Empty:
                break
        self.window.after(100, self.check_log_queue)

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = MCQSplitterGUI()
    app.run()