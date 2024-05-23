import serial
import time

def open_serial(port, baud_rate):
    try:
        ser = serial.Serial(port, baud_rate)
        time.sleep(2)  # Allow time for the connection to initialize
        print(f"Connected to the Arduino on {port}.")
        return ser
    except serial.SerialException as e:
        print(f"Error opening serial port {port}: {e}")
        return None

def send_gcode_command(ser, command):
    full_command = f"{command}\n"
    ser.write(full_command.encode('utf-8'))
    ser.flush()
    print(f"Sent: {command}")
    response = ser.readline().decode('utf-8').strip()
    print(f"Received response: {response}")

def main_menu():
    print("\nArduino CNC Controller")
    print("1 - Move X-axis motor to the right")
    print("2 - Move X-axis motor to the left")
    print("3 - Move Y-axis motor forward")
    print("4 - Move Y-axis motor backward")
    print("5 - Quit")
    return input("Choose an option: ")

def perform_action(ser, choice):
    if choice == '1':
        send_gcode_command(ser, "G91")  # Relative positioning
        send_gcode_command(ser, "G0 X10")  # Move X-axis motor 10 units to the right
    elif choice == '2':
        send_gcode_command(ser, "G91")  # Relative positioning
        send_gcode_command(ser, "G0 X-10")  # Move X-axis motor 10 units to the left
    elif choice == '3':
        send_gcode_command(ser, "G91")  # Relative positioning
        send_gcode_command(ser, "G0 Y10 Z10")  # Move X-axis motor 10 units to the right
    elif choice == '4':
        send_gcode_command(ser, "G91")  # Relative positioning
        send_gcode_command(ser, "G0 Y-10 Z-10")  # Move X-axis motor 10 units to the right
    elif choice == '5':
        print("Exiting program.")
    else:
        print("Invalid option. Please try again.")

if __name__ == "__main__":
    port = input("Enter the Arduino serial port (e.g., COM3 or /dev/ttyACM0): ") #/dev/cu.usbserial-10
    baud_rate = 115200
    ser = open_serial(port, baud_rate)
    if ser:
        while True:
            user_choice = main_menu()
            if user_choice == '5':
                break
            perform_action(ser, user_choice)
        ser.close()  # Close the serial connection when done
