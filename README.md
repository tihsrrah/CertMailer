# CertMailer
A local python script for automated mailing of certificates
Certificate Generator & Mailer

A simple Python GUI tool to automatically generate participation certificates and email them to participants.
Created for college events, fests, and societies that need to send certificates to many people quickly.

What this tool does

Lets you load a certificate template (JPG/PNG)

Reads a participant list (Excel or CSV) with names & emails

Prints each participant’s name on the certificate in Poppins Bold

Adds a unique security code like quiz-25-010

Saves certificates as PDFs

Can optionally email each certificate through Gmail

Shows a live progress log for every step

How to use it

Run the Python script:

python cert_mailer.py


Use the GUI to select:

Certificate template image

Participant list

Poppins-Bold.ttf and Arial.ttf

Output folder

Enter:

Event code (e.g., quiz)

Year (e.g., 25)

Sender Gmail ID & password

Choose:

Generate Certificates Only

Generate & Send Certificates

Sit back. The certificates will be created (and emailed if you selected that option).

Participant file format

Your Excel/CSV needs only:

Name	Email

The row number decides the serial number for the security code.

Security code format

Each certificate gets a verification code:

<eventcode>-<year>-<serial>


Example:

quiz-25-010


Printed bottom-left in small light-gray Arial text.

Requirements

Install these Python packages:

pip install pillow pandas openpyxl

Converting to EXE (optional)
pyinstaller --onefile --noconsole cert_mailer.py

Notes

This tool works with any certificate design as long as the template resolution stays the same.

Gmail password is used only for sending – it’s not saved anywhere.

Best practice: use a separate Gmail account for certificate sending.
