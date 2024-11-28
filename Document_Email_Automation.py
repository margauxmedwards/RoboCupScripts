'''
Document_Email_Automation.py
This script automates the generation of documents and sending of emails to recipients for RoboCup Junior Australia.
It reads participant details from a CSV file, splits a multi-page PDF into individual documents, and names each document
according to the team or individual's name. The script then sends emails to recipients with the generated documents attached.
The email body is personalized with the recipient's name and includes a custom message. The script uses the
Outlook application for sending emails.
Classes:
    DocumentGenerator: Responsible for generating individual documents and updating the CSV file with the file paths.
    EmailSender: Responsible for creating email drafts and sending emails to recipients with the generated documents attached.
Usage:
    - Ensure that the required libraries (os, csv, pandas, PyPDF2, win32com) are installed before running the script.
    - Update the file paths and email details in the main section of the script as needed.
    - Run the script to generate documents and send emails.
Author: Margaux Edwards
Date: 28/11/2024
Email: margaux.edwards@robocupjunior.org.au
'''

# CSV file format:
# Team Name,Organisation,Division,Award,Mentor_Name,Mentor_Email

import os
import csv
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import win32com.client as win32


class DocumentGenerator:
    def __init__(self, pdf_path, csv_file, output_folder, document_type):
        """
        Initialize the DocumentGenerator with the paths and document type.

        Args:
            pdf_path (str): Path to the input PDF file.
            csv_file (str): Path to the CSV file containing participant details.
            output_folder (str): Directory to save the generated Documents.
            document_type (str): Type of document (e.g., 'Award', 'Participation').
        """
        self.pdf_path = pdf_path
        self.csv_file = csv_file
        self.output_folder = output_folder
        self.document_type = document_type

    def generate_Documents(self):
        """
        Generate Documents by splitting a multi-page PDF and naming each page
        according to the Name column in the provided CSV file. Adds the output
        file path to the CSV for future reference.
        """
        # Ensure the output folder exists or create it
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
            print(f"Output folder created: {self.output_folder}")
        else:
            print(f"Output folder already exists: {self.output_folder}")

        # Read the CSV file and extract the data
        with open(self.csv_file, "r") as file:
            reader = csv.DictReader(file)
            rows = list(reader)  # Convert to a list of dictionaries
            fieldnames = reader.fieldnames + ["File Path"]  # Add a new "File Path" column

        # Load the PDF
        reader = PdfReader(self.pdf_path)
        total_pages = len(reader.pages)

        # Ensure the number of rows matches the number of pages
        if len(rows) != total_pages:
            raise ValueError("The number of entries in the CSV doesn't match the number of pages in the PDF.")

        # Split the PDF and name each page with the corresponding Name from the CSV
        for i, row in enumerate(rows):
            writer = PdfWriter()
            writer.add_page(reader.pages[i])

            # Create the output path using the Name column and document type
            name = row["Team Name"]
            output_file_name = f"{name}_{self.document_type}.pdf"
            output_path = os.path.join(self.output_folder, output_file_name)
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            # Add the file path to the current row
            row["File Path"] = output_path
            print(f"Created: {output_path}")

        # Write the updated rows with the new "File Path" column back to the CSV
        updated_csv_file = os.path.join(self.output_folder, f"{self.document_type}_Documents_Updated.csv")
        with open(updated_csv_file, "w", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()  # Write the header
            writer.writerows(rows)  # Write all rows with the updated data

        print(f"All Documents generated successfully. Updated CSV saved at: {updated_csv_file}")
        return updated_csv_file


class EmailSender:
    def __init__(self, csv_file, email_subject, email_body_template, sender_name, sender_title, organisation):
        """
        Initialize the EmailSender with the email details and recipient data.

        Args:
            csv_file (str): Path to the CSV file containing recipient data.
            email_subject (str): Subject of the email.
            email_body_template (str): Template for the email body.
            sender_name (str): Name of the sender.
            sender_title (str): Title of the sender.
            organisation (str): organisation name.
        """
        self.data = pd.read_csv(csv_file)
        self.email_subject = email_subject
        self.email_body_template = email_body_template
        self.sender_name = sender_name
        self.sender_title = sender_title
        self.organisation = organisation

        # Ensure proper file paths
        self.data['File Path'] = self.data['File Path'].apply(lambda x: x.replace('\\', '/'))

    def create_drafts(self, sample_size=None):
        """
        Create email drafts for the given data.

        Args:
            sample_size (int, optional): Number of random rows to use for draft testing. If None, drafts are created for all rows.
        """
        outlook = win32.Dispatch('outlook.application')

        if sample_size:
            data = self.data.sample(n=sample_size)  # Randomly sample rows for testing
        else:
            data = self.data

        for _, row in data.iterrows():
            # Create email
            mail = outlook.CreateItem(0)
            mail.To = row['Mentor_Email']
            mail.Subject = self.email_subject

            # Generate the email body by replacing placeholders with actual data
            email_body = self.email_body_template.format(
                mentor_name=row['Mentor_Name'],
                division=row['Division'],
                sender_name=self.sender_name,
                sender_title=self.sender_title,
                organisation=self.organisation
            )
            mail.Body = email_body

            filepath = row['File Path']
            filepath = os.path.abspath(filepath)  # Convert to an absolute path for safety
            # Attach the PDF
            mail.Attachments.Add(filepath)

            # Save the email as a draft
            mail.Save()
            print(f"Draft created for: {row['Mentor_Email']}")

        print("Draft emails created successfully! Check your Outlook Drafts folder.")

    def send_emails(self):
        """
        Send emails in bulk based on the given data.
        """
        outlook = win32.Dispatch('outlook.application')

        for _, row in self.data.iterrows():
            # Create email
            mail = outlook.CreateItem(0)
            mail.To = row['Mentor_Email']
            mail.Subject = self.email_subject

            # Generate the email body by replacing placeholders with actual data
            email_body = self.email_body_template.format(
                mentor_name=row['Mentor_Name'],
                division=row['Division'],
                sender_name=self.sender_name,
                sender_title=self.sender_title,
                organisation=self.organisation
            )
            mail.Body = email_body

            filepath = row['File Path']
            filepath = os.path.abspath(filepath)  # Convert to an absolute path for safety
            # Attach the PDF
            mail.Attachments.Add(filepath)

            # Send the email
            mail.Send()
            print(f"Email sent to: {row['Mentor_Email']}")

        print("All emails sent successfully!")


if __name__ == "__main__":
    # Document Generation
    pdf_path = r"XXX.pdf"  # Replace with your input PDF file path. Ths PDF should contain all the documents to be generated in a single file. This can easily be a bulk generated file from Canva.
    csv_file = r"XXX.csv"  # Replace with your input CSV file path. All the details from the registration system
    output_folder = r"Document_Output_Directory"  # Replace with your desired output folder
    document_type = "XXX"  # Replace with the desired document type (e.g., 'Participation', 'Volunteer', 'Award')

    cert_generator = DocumentGenerator(pdf_path, csv_file, output_folder, document_type)
    updated_csv = cert_generator.generate_Documents()
    
    sender_name = "XXX TEMPLATE SENDER NAME XXX"  # Name of the sender
    sender_title = "XXX TEMPLATE SENDER TITLE XXX"  # Title of the sender
    organisation = "XXX TEMPLATE ORGANISATION XXX"  # Name of the organisation

    # Email Sending
    email_subject = "XXX TEMPLATE SUBJECT XXX"  # Subject of the email
    email_body_template = """Dear {mentor_name},
    
XXX TEMPLATE MESSAGE XXX
Example: {division}
XXX TEMPLATE MESSAGE XXX

{sender_name}
{sender_title}
{organisation}
"""

    email_sender = EmailSender(updated_csv, email_subject, email_body_template, sender_name, sender_title, organisation)

    # Create Drafts (Optional)
    email_sender.create_drafts(sample_size=3)  # Uncomment to create drafts for testing

    # Send Emails (Uncomment to send emails in bulk)
    email_sender.send_emails()
