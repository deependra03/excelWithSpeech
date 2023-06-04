import speech_recognition as sr
import openpyxl

# Create a new Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active

# Initialize the recognizer
recognizer = sr.Recognizer()


# Function to convert speech to text
def convert_speech_to_text():
    with sr.Microphone() as source:
        print("Speak something...")
        audio = recognizer.listen(source)

        try:
            # Recognize speech using the Google Speech Recognition API
            text = recognizer.recognize_google(audio)
            return text
        except sr.UnknownValueError:
            print("Unable to recognize speech")
        except sr.RequestError as e:
            print("Error occurred; {0}".format(e))


# Prompt user for speech input to fill Excel rows
field_names = ['Name', 'Age', 'Email']  # Example field names
# Set header names
header_names = field_names
num_rows = 3  # Number of rows to fill
for row in range(1, num_rows + 1):
    for i, field in enumerate(field_names, start=1):
        print(f"Speak the value for {field} in row {row}:")
        value = convert_speech_to_text()
        sheet.cell(row=row, column=i, value=value)

# Save the Excel file
workbook.save('speech_to_excel.xlsx')
print("Speech input saved in speech_to_excel.xlsx")
