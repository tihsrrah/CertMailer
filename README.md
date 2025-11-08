# CertMailer
A simple Python GUI tool to automatically generate participation
certificates and send them via email.

## Features

-   Load a certificate template (PNG/JPG)
-   Read participant list (Excel/CSV)
-   Print participant names in Poppins Bold
-   Add unique security codes like `quiz-25-010`
-   Save certificates as PDFs
-   Optional Gmail-based auto‑mailing
-   Live progress log

## How to Use

1.  Run:

        python cert_mailer.py

2.  In the GUI:

    -   Select certificate template
    -   Select participant Excel/CSV
    -   Select Poppins-Bold.ttf & Arial.ttf
    -   Choose output folder

3.  Enter:

    -   Event code (e.g., quiz)
    -   Year (e.g., 25)
    -   Sender Gmail ID, Email subject and app password

4.  Choose:

    -   Generate certificates only\
    -   Generate & send certificates

## Participant File Format

  Name   Email
  ------ -------

Row number becomes the serial for the security code.

## Security Code Format

    <eventcode>-<year>-<serial>

Example:

    quiz-25-010

## Requirements

    pip install pillow pandas openpyxl



## Notes

-   Works with any template as long as resolution stays same.
-   Black underline bar is needed to recognise where the name goes.
-   Gmail password not saved.
-   Recommended: use a separate Gmail account for sending.

