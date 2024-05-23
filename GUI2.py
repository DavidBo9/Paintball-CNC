import sys
import serial
import openpyxl
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QWidget, QComboBox, QFileDialog, QMessageBox)
from openpyxl.styles import PatternFill

class IndustrialCncGui(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.ser = self.init_serial('/dev/cu.usbserial-10', 115200)  # Update the port according to your setup
        self.init_ui()

    def init_serial(self, port, baud_rate):
        try:
            ser = serial.Serial(port, baud_rate)
            print(f"Connected to the Arduino on {port}.")
            return ser
        except Exception as e:
            QMessageBox.critical(self, "Serial Connection Error", f"Failed to connect to the Arduino: {e}")
            return None

    def init_ui(self):
        # Setup UI elements
        self.setWindowTitle('Industrial Paintball CNC Controller')
        self.setGeometry(100, 100, 1280, 720)
        self.setup_ui_elements()
        self.style_interface()

    def setup_ui_elements(self):
        # Create main layout
        main_layout = QHBoxLayout()
        self.create_control_panel(main_layout)
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def create_control_panel(self, main_layout):
        left_layout, right_layout = QVBoxLayout(), QVBoxLayout()

        # Left Layout: Excel File Processing
        self.load_button = QPushButton('Load Excel File')
        self.load_button.clicked.connect(self.load_excel_file)
        self.layer_combobox = QComboBox()
        self.layer_preview = QLabel('Layer Preview')
        self.process_button = QPushButton('Process Layer')
        left_layout.addWidget(self.load_button)
        left_layout.addWidget(self.layer_combobox)
        left_layout.addWidget(self.layer_preview)
        left_layout.addWidget(self.process_button)

        # Right Layout: CNC Controls
        self.add_movement_controls(right_layout)

        # Add left and right layouts to the main layout
        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)

    def add_movement_controls(self, layout):
        # Mapping GUI buttons to corresponding serial commands for CNC movements
        move_commands = {
            'UP': 'G0 Y10',
            'DOWN': 'G0 Y-10',
            'LEFT': 'G0 X-10',
            'RIGHT': 'G0 X10'
        }
        
        for direction, command in move_commands.items():
            button = QPushButton(direction, self)
            button.clicked.connect(lambda _, c=command: self.send_gcode_command(c))
            layout.addWidget(button)

    def send_gcode_command(self, command):
        # Sends the G-code command to the Arduino through serial
        if self.ser and self.ser.isOpen():
            self.ser.write(f"{command}\n".encode('utf-8'))
            self.ser.flush()
            print(f"Sent: {command}")
            # This assumes Arduino sends back a simple OK response for every command sent
            response = self.ser.readline().decode('utf-8').strip()
            print(f"Received response: {response}")
        else:
            print("Serial connection is not open.")

    def load_excel_file(self):
        # Implementation for loading and processing Excel files remains unchanged
        pass

    def process_layer(self):
        # Implementation for processing layer remains unchanged
        pass

    def style_interface(self):
        # Styling for the interface; implementation remains unchanged
        pass

def main():
    app = QApplication(sys.argv)
    gui = IndustrialCncGui()
    gui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
