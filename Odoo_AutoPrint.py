import os
import time
import win32print
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileMovedEvent
import subprocess
from PyPDF2 import PdfReader, PdfWriter

class Watcher:
    DIRECTORY_TO_WATCH = "C:\\Users\\MathieuBourassa.DESKTOP-04JCUJM\\Downloads"

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.DIRECTORY_TO_WATCH, recursive=True)
        self.observer.start()
        print("Monitoring for new PDF files...")
        try:
            while True:
                time.sleep(5)
        except KeyboardInterrupt:
            self.observer.stop()
            print("Observer Stopped")

        self.observer.join()


class Handler(FileSystemEventHandler):
    @staticmethod
    def on_moved(event):
        if isinstance(event, FileMovedEvent):
            new_file_name = os.path.basename(event.dest_path)
            print(f"File moved/renamed: {new_file_name}")  # Debug print

            if new_file_name.endswith(".pdf") and new_file_name.startswith("Product Label (PDF)"):
                print(f"New PDF file detected: {new_file_name}")
                print_file(event.dest_path)


def print_file(file_path):
    printer_name = win32print.GetDefaultPrinter()
    print(f"Printing on: {printer_name}")

    try:
        # Use the default application for PDFs to print the file
        os.startfile(file_path, "print")
        print(f"Sent to printer: {file_path}")

        # Wait for a moment to ensure the file has been sent to the printer
        time.sleep(2)  # Adjust this as needed

        # Delete the file
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"Deleted file: {file_path}")
        else:
            print("Error: File not found for deletion.")
    except Exception as e:
        print(f"Error printing file: {e}")

if __name__ == '__main__':
    w = Watcher()
    w.run()
