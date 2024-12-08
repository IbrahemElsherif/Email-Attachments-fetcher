import win32com.client
import os
import pikepdf
import pdfplumber


class EmailAttachmentFetch:
    """
    A class to fetch email attachments from Outlook based on a specific subject.
    """
    def __init__(self, subject, output_dir, pdfs_only=False):
        """
        Initialize EmailAttachmentFetch with given parameters.

        Args:
            subject (str): Email subject to search
            output_dir (str): Directory to save attachments
            pdfs_only (bool, optional): Extract only PDF files. Defaults to False.
        """
        self.subject = subject
        self.output_dir = output_dir
        self.pdfs_only = pdfs_only

    def fetch_attachments(self):
        """
        Fetch attachments from emails matching the specified subject.

        Returns:
            list: Paths of extracted attachment files
        """
        # Connect to Outlook application
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        except Exception as e:
            print(f"Error Accessing {outlook}: {e}")
            return []

        inbox = outlook.GetDefaultFolder(6)  # Inbox folder
        # inbox = outlook.Folders["Inbox"]
        # Create output folder if it doesn't exist
        os.makedirs(self.output_dir, exist_ok=True)

        messages = inbox.Items
        # Search for emails with the specified subject
        messages = messages.Restrict(f"[Subject] = '{self.subject}'")
        print(f"Found {messages.Count} emails with subject: {self.subject}")

        extracted_files = []

        for message in messages:
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    file_name = attachment.FileName
                    # Check file extension for PDF filtering
                    if self.pdfs_only and not file_name.lower().endswith('.pdf'):
                        continue

                    # Save attachment
                    attachment_path = os.path.join(self.output_dir, file_name)
                    attachment.SaveAsFile(attachment_path)
                    extracted_files.append(attachment_path)
                    print(f"Saved: {attachment_path}")

        return extracted_files


class PdfProcessing:
    """
    A class to process PDF files with various operations.
    """

    def __init__(self, pdf_files, output_dir):
        """
        Initialize PdfProcessing with given PDF files and output directory.

        Args:
            pdf_files (list): List of PDF file paths
            output_dir (str): Directory to save processed PDFs
        """
        self.pdf_files = [f for f in pdf_files if f.lower().endswith('.pdf')]
        self.output_dir = output_dir

    def combine_pdf(self):
        """
        Combine multiple PDF files into a single PDF.
        Saves the combined PDF in the output directory.
        """
        combined_pdf = os.path.join(self.output_dir, "Combined_documents.pdf")

        try:
            pdf = pikepdf.Pdf.new()

            for pdf_file in self.pdf_files:
                src = pikepdf.Pdf.open(pdf_file)
                pdf.pages.extend(src.pages)
            pdf.save(combined_pdf)
            print(f"Combined/Merged PDFs saved to {combined_pdf}")
        except Exception as e:
            print(f"Error combining PDFs: {e}")

    def extract_pages_with_keyword(self, keyword):
        """
        Extract pages containing a specific keyword from PDF files.
        Args:
             keyword (str): Keyword to search for in PDF pages
        """
        for pdf_file in self.pdf_files:
            output_pdf = pikepdf.Pdf.new()

            try:
                with pdfplumber.open(pdf_file) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        # Extract text from the page
                        page_text = page.extract_text() or ''

                        if keyword.lower() in page_text.lower():
                            # If keyword found, add the corresponding page from the original PDF
                            src_pdf = pikepdf.Pdf.open(pdf_file)
                            output_pdf.pages.append(src_pdf.pages[page_num])

                # If pages with the keyword were found, save them to a new PDF file
                if len(output_pdf.pages) > 0:
                    output_file = os.path.join(self.output_dir,
                                               f"{os.path.splitext(os.path.basename(pdf_file))[0]}_keyword_pages.pdf")
                    output_pdf.save(output_file)
                    print(f"Pages with keyword '{keyword}' saved to {output_file}")
                else:
                    print(f"No pages containing keyword '{keyword}' found in {pdf_file}")

            except Exception as e:
                print(f"Error processing {pdf_file}: {e}")

    def find_word_freq(self, word):
        """
        Find frequency of a specific word across PDF files.
        Args:
             word (str): Word to count occurrences of
        """
        for pdf_file in self.pdf_files:
            word_count = 0
            try:
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        # Extract text from the page
                        page_text = page.extract_text() or ''
                        word_count += page_text.lower().count(word.lower())

                print(f"'{word}' found {word_count} times in {pdf_file}")
            except Exception as e:
                print(f"Error processing {pdf_file}: {e}")

if __name__ == "__main__":
    subject = input("Enter the email subject: ")
    output_dir = input("Enter the path for output folder: ") or "C:\\Attachments"

    # Extract PDFs based on user input
    extract_pdf = input("Do you want to extract only PDFs? (yes/no): ").strip().lower() == "yes"

    # Fetch attachments
    fetch = EmailAttachmentFetch(subject, output_dir, pdfs_only=extract_pdf)
    attachments = fetch.fetch_attachments()

    # Validate attachments
    if not attachments:
        print("No attachments found.")
        exit()

    # Process PDFs
    pdf_processor = PdfProcessing(attachments, output_dir)

    if input("Do you want to combine the PDFs? (yes/no): ").strip().lower() == "yes":
        pdf_processor.combine_pdf()

    if input("Do you want to extract pages with a keyword? (yes/no): ").strip().lower() == "yes":
        keyword = input("Enter the keyword: ")
        pdf_processor.extract_pages_with_keyword(keyword)

    if input("Do you want to find word frequency? (yes/no): ").strip().lower() == "yes":
        word = input("Enter the word to search: ")
        pdf_processor.find_word_freq(word)
