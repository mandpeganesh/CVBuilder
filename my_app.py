"""
CV Builder Application

This script creates a CV (Curriculum Vitae) document using user input.
It uses the python-docx library to create a Word document and pyttsx3 for text-to-speech functionality.
"""
from docx import Document
from docx.shared import Inches
import pyttsx3
import os


def speak(text):
    """
    Use text-to-speech to speak the given text.

    Args:
        text (str): The text to be spoken.
    """
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()


# Create a new Document
document = Document()

# Add profile picture if available
if os.path.exists('me.jpg'):
    document.add_picture('me.jpg', width=Inches(2.0))
else:
    print('Profile picture not found. Skipping.')

# Collect personal information
name = input('What is your name? ')
speak('Hello ' + name + ' how are you today?')

speak('What is your phone number?')
phone_number = input('What is your phone number? ')
email = input('What is your email? ')

document.add_paragraph(name + ' | ' + phone_number + ' | ' + email)

# Add 'About Me' section
document.add_heading('About Me')
about_me = input('Tell me about yourself: ')
document.add_paragraph(about_me)

# Function to add work experience
def add_experience():
    company = input('Enter company: ')
    from_date = input('Enter start date: ')
    to_date = input('Enter end date: ')
    experience_details = input('Describe your experience at ' + company + ': ')
    
    p = document.add_paragraph()
    p.add_run(company + '  ').bold = True
    p.add_run(from_date + '-' + to_date + '\n').italic = True
    p.add_run(experience_details)

# Add work experience section
document.add_heading('Work Experience')

while True:
    add_experience()
    has_more_experiences = input('Do you have more experiences? Yes or No: ')
    if has_more_experiences.lower() != 'yes':
        break

# Function to add skills
def add_skill():
    skill = input('Enter skill: ')
    p = document.add_paragraph(skill)
    p.style = 'List Bullet'

# Add skills section
document.add_heading('Skills')

while True:
    add_skill()
    has_more_skills = input('Do you have more skills? Yes or No: ')
    if has_more_skills.lower() != 'yes':
        break

# Footer section
section = document.sections[0]
footer = section.footer
p = footer.paragraphs[0]
p.text = 'CV generated using GNM\'s CV generator'

# Save the document
document_name = input('Enter the name for your CV file (e.g., my_cv.docx): ')
document.save(document_name)
