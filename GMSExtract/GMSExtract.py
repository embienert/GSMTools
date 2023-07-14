"""
Script created by Martin Bienert.
Python version used: 3.8.5 (should be compatible with any 3.x, as no 'new' features were used)

This script is intended to be used for extracting the H-/P-/EUH-Statements
from chemicals' safety data sheets. The extracted Statements will be returned
as TAB-separated Lists of the COMMA-separated statements in order of their discovery
within the input.
e.g. H319, H335, H315	P261, P302 + P352, P280, P305 + P351 + P338, P271	EUH061

Input may be the path to a pdf file, finished with an empty input line.
If a path to a pdf file was recognized the program will confirm, that the input will be interpreted as such.
    > aceton.pdf
    >

    Interpreting input as path of pdf file.

Input may alternatively be the entire safety data sheet as text, finished with an empty input line.
    > Sicherheitsinformationen gemäß GHS Gefahrensymbol(e)
    > Gefahrenhinweis(e)
    > H315: Verursacht Hautreizungen.
    > H319: Verursacht schwere Augenreizung.
    > H335: Kann die Atemwege reizen.
    > Sicherheitshinweis(e)
    > P261: Einatmen von Staub/ Rauch/ Gas/ Nebel/ Dampf/ Aerosol vermeiden.
    > P271: Nur im Freien oder in gut belüfteten Räumen verwenden.
    > P280: Schutzhandschuhe/ Augenschutz/ Gesichtsschutz tragen.
    > P302 + P352: BEI BERÜHRUNG MIT DER HAUT: Mit viel Wasser waschen.
    > P305 + P351 + P338: BEI KONTAKT MIT DEN AUGEN: Einige Minuten lang behutsam mit Wasser spülen.
    > Eventuell vorhandene Kontaktlinsen nach Möglichkeit entfernen. Weiter spülen.
    > Ergänzende Gefahrenhinweise EUH061
    > SignalwortAchtungLagerklasse10 - 13 Sonstige Flüssigkeiten und
    > FeststoffeWGKWGK 1 schwach wassergefährdendEntsorgung3
    >

"""

from pdfminer.pdfdocument import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams
from multiprocessing import Pool
from typing import List, Tuple
from io import StringIO
from time import time
from glob import glob
import sys
import re

out_file = None


def timeit(func):
    def inner_timeit(*args, **kwargs):
        start = time()
        return_value = func(*args, **kwargs)
        print(f"[TIMEIT {func.__name__}]: {time() - start}s")

        return return_value

    return inner_timeit


class GMSExtract:
    h_pattern = re.compile(r"(?:(?<!EU)H[2-4][0-9]{2}[dDfF]{0,2}(?![0-9]))"
                           r"(?:\s*\+\s*(?<!EU)H[2-4][0-9]{2}[dDfF]{0,2}(?![0-9]))*")
    p_pattern = re.compile(r"(?:P[1-5][0-9]{2}(?![0-9]))(?:\s*\+\s*P[1-5][0-9]{2}(?![0-9]))*")
    euh_pattern = re.compile(r"(?:EUH[0-9]{3}[dDfF]?(?![0-9]))(?:\s*\+\s*EUH[0-9]{3}[dDfF]?(?![0-9]))*")
    wgk_pattern = re.compile(r"WGK.*?[0-3](?![0-9])")
    WGK_pattern = re.compile(r"[Ww]assergefährdungsklasse.*?[0-3](?![0-9])")
    cas_pattern = re.compile(r"[0-9]{1,8}-[0-9]{2}-[0-9](?![0-9])")

    OUTPUT_SEP = "\t"

    USE_MULTIPROCESSING = True

    # @timeit
    @staticmethod
    def read_pdf(filename: str) -> str:
        """
        Code taken from @RattleyCooper and @Trenton McKinney at
        https://stackoverflow.com/questions/26494211/extracting-text-from-a-pdf-file-using-pdfminer-in-python

        Read content of pdf file to string

        :param filename: relative or absolute path of input pdf file
        :return: content of input pdf file as text
        """

        resource_manager = PDFResourceManager()
        params = LAParams()
        output_string = StringIO()
        codec = 'utf-8'
        device = TextConverter(resource_manager, output_string, codec=codec, laparams=params)
        interpreter = PDFPageInterpreter(resource_manager, device)

        fp = open(filename, 'rb')
        page_numbers = set()
        caching = True
        password = ""
        max_pages = 12

        try:

            for page in PDFPage.get_pages(fp, page_numbers, maxpages=max_pages, password=password, caching=caching,
                                          check_extractable=True):
                interpreter.process_page(page)
        except PDFTextExtractionNotAllowed:
            # print("Could not read " + filename + " : File is protected.")
            return ""

        content_text = output_string.getvalue()

        fp.close()
        device.close()
        output_string.close()

        return content_text

    @staticmethod
    # @timeit
    def read_pdf_multiple(filenames: List[str]) -> List[str]:
        """
        Read contents of all files in filenames

        :param filenames: List containing filenames and paths to pdf files
        :return: List containing the contents of each of the passed pdf files as strings
        """
        file_contents: List[str] = []

        if GMSExtract.USE_MULTIPROCESSING:
            with Pool(len(filenames)) as executor:
                file_contents = executor.map(GMSExtract.read_pdf, filenames)

            return file_contents

        for filename in filenames:
            file_contents.append(GMSExtract.read_pdf(filename))

        return file_contents

    @staticmethod
    # @timeit
    def normalize_string(string: str) -> str:
        """
        Replace all linebreaks with whitespaces and normalize amount of whitespaces around '+' to exactly one

        :param string: string that is to be normalized
        :return: normalized string
        """
        return re.sub(r"\n+", " ", re.sub(r"\s*\+\s*", " + ", string))

    @staticmethod
    def match_h(string: str) -> List[str]:
        """
        Find all matches for H-Statements in the given input string using regular expressions

        :param string: input string containing H-Statements
        :return: list of unique H-Statements as strings
        """
        return sorted(list(set(GMSExtract.h_pattern.findall(string))))

    @staticmethod
    def match_p(string: str) -> List[str]:
        """
        Find all matches for P-Statements in the given input string using regular expressions

        :param string: input string containing P-Statements
        :return: list of unique P-Statements as strings
        """
        return sorted(list(set(GMSExtract.p_pattern.findall(string))))

    @staticmethod
    def match_euh(string: str) -> List[str]:
        """
        Find all matches for EUH-Statements in the given input string using regular expressions

        :param string: input string containing EUH-Statements
        :return: list of unique EUH-Statements as strings
        """
        return sorted(list(set(GMSExtract.euh_pattern.findall(string))))

    @staticmethod
    def match_wgk(string: str) -> str:
        """
        Find first match for WGK (Wassergefährdungsklasse) in the given input string using regular expressions

        :param string: input string containing WGK information
        :return: WGK as string, empty if none found
        """
        match: List[str] = GMSExtract.wgk_pattern.findall(string) + GMSExtract.WGK_pattern.findall(string)
        return "" if match == [] else match[0][-1]

    @staticmethod
    def match_cas(string: str) -> List[str]:
        """
        Find all matches for CAS-Numbers in the given input string using regular expressions

        :param string: input string containing CAS-Numbers
        :return: list of unique CAS-Numbers as strings
        """
        match: List[str] = GMSExtract.cas_pattern.findall(string)
        return sorted(list(set(filter(GMSExtract.is_cas_valid, match))))

    @staticmethod
    def is_cas_valid(cas: str) -> bool:
        """
        Check if the given CAS-Number is valid by calculating and checking the checksum according to
        https://de.wikipedia.org/wiki/CAS-Nummer

        :param cas: CAS-Number as string
        :return: whether or not the given CAS-Number is valid
        """
        cas_split = cas.split("-")
        numbers, checksum = "".join(cas_split[:-1]), int(cas_split[-1])

        cas_sum = 0
        for value, number in enumerate(reversed(numbers)):
            cas_sum += (value + 1) * int(number)

        return (cas_sum % 10) == checksum

    @staticmethod
    # @timeit
    def process(string: str) -> Tuple[List[str], List[str], List[str], List[str], str]:
        """
        Find all matches for H-/P-/EUH-Statements and WGK (Wassergefährdungsklasse) in the given input string.

        :param string: input string containing H-/P-/EUH-Statements
        :return: three lists containing the H-/P-/EUH-Statements as described in the designated methods and WGK
        """

        if string == "":
            h_match = ["Could not read. File is protected."]
            return [], h_match, [], [], ""

        normalized_string: str = GMSExtract.normalize_string(string)

        h_match = GMSExtract.match_h(normalized_string)
        p_match = GMSExtract.match_p(normalized_string)
        euh_match = GMSExtract.match_euh(normalized_string)
        wgk_match = GMSExtract.match_wgk(normalized_string)
        cas_match = GMSExtract.match_cas(normalized_string)

        return cas_match, h_match, p_match, euh_match, wgk_match

    @staticmethod
    # @timeit
    def process_all(strings: List[str]) -> Tuple[List[List[str]], List[List[str]], List[List[str]], List[List[str]],
                                                 List[str]]:
        """
        Process each string in the passed list of strings as described in the process method

        :param strings: list of input strings containing H-/P-/EUH-Statements, WGK and CAS-Number
        :return: three lists containing the lists of H-/P-/EUH-Statements from each file and a list of WGKs
        """

        matches: Tuple[List[List[str]], List[List[str]], List[List[str]], List[List[str]], List[str]] = \
            ([], [], [], [], [])

        for string in strings:
            for index, match in enumerate(GMSExtract.process(string)):
                matches[index].append(match)

        return matches

    @staticmethod
    def print_excel(cas_match: List[str], h_match: List[str], p_match: List[str], euh_match: List[str], wgk: str,
                    filename: str) -> None:
        """
        Print the H-/P-/EUH-Statements, the WGK (Wassergefährdungsklasse) and CAS-Numbers in a manner that allow the
        output string to be copy + pasted into excel

        :param cas_match: List containing all matches for CAS-Numbers as described in match_cas
        :param h_match: List containing all matches for H-Statements as described in match_h
        :param p_match: List containing all matches for P-Statements as described in match_p
        :param euh_match: List containing all matches for EUH-Statements as described in match_euh
        :param wgk: String of WGK ("Wassergefährdungsklasse") value, empty if non found
        :param filename: Name of the file the H-/P-/EUH-Statements were taken from. Empty if not from file
        """

        if len(cas_match) + len(h_match) + len(p_match) + len(euh_match) + len(wgk) == 0:
            print(f"\nFinished extracting. No Statements were found{(' in ' + filename) if filename != '' else ''}.")
            return

        prefix = filename + ": \t" if filename != "" else ""
        print("\nFinished extracting. The following line can be copy + pasted into excel.")
        print(prefix + GMSExtract.OUTPUT_SEP.join([", ".join(cas_match), ", ".join(h_match), ", ".join(p_match),
                                                   ", ".join(euh_match), wgk]))

    @staticmethod
    def string_excel(cas_match: List[str], h_match: List[str], p_match: List[str], euh_match: List[str],
                     wgk: str, filename: str) -> Tuple[str, bool]:
        """
        Create string from the H-/P-/EUH-Statements, the WGK (Wassergefährdungsklasse) and CAS-Numbers in a manner that
        allow the returned string to be copy + pasted into excel

        :param cas_match: List containing all matches for CAS-Numbers as described in match_cas
        :param h_match: List containing all matches for H-Statements as described in match_h
        :param p_match: List containing all matches for P-Statements as described in match_p
        :param euh_match: List containing all matches for EUH-Statements as described in match_euh
        :param wgk: String of WGK ("Wassergefährdungsklasse") value, empty if non found
        :param filename: Name of the file the H-/P-/EUH-Statements were taken from. Empty if not from file
        :return: formatted string containing the H-/P-/EUH-Statements and the WGK
        """
        global out_file

        prefix = filename + "\t" if filename != "" else ""

        if len(h_match) + len(p_match) + len(euh_match) + len(wgk) == 0:
            return_string = prefix + "No Statements found."
            out_file.write(return_string.encode('utf-8', 'ignore') + b"\n")

            return return_string, False

        return_string = prefix + GMSExtract.OUTPUT_SEP.join([", ".join(cas_match), ", ".join(h_match),
                                                             ", ".join(p_match), ", ".join(euh_match), wgk])
        out_file.write(return_string.encode('utf-8', 'ignore') + b"\n")

        return return_string, True

    @staticmethod
    # @timeit
    def string_excel_all(cas_matches: List[List[str]], h_matches: List[List[str]], p_matches: List[List[str]],
                         euh_matches: List[List[str]], wgks: List[str], filenames: List[str]) -> str:
        """
        Create string from the H-/P-/EUH-Statements, the WGK and CAS-Numbers from each file as described in the
        string_excel method and concatenate them to a table-like output string.

        :param cas_matches: List containing all lists of matches for CAS-Numbers for each input file/text
        :param h_matches: List containing all lists of matches for H-Statements for each input file/text
        :param p_matches: List containing all lists of matches for P-Statements for each input file/text
        :param euh_matches: List containing all lists of matches for EUH-Statements for each input file/text
        :param wgks: List containing all WGKs from each input file/text
        :param filenames: List containing all input filenames.
        :return: Concatenated table-like string containing formatted data on all input files/texts
        """
        found_statements: List[str] = []

        for values in zip(cas_matches, h_matches, p_matches, euh_matches, wgks, filenames):
            string, found = GMSExtract.string_excel(*values)
            found_statements.append(string)

        return "\n".join(found_statements)


def get_input() -> Tuple[List[str], List[str]]:
    """
    Read text from input prompt until an empty line is sent and the input up to this point is non-empty.
    If the text input is recognized to be a path to a pdf file, there will be an attempt to open the file
    and return its contents as text. If the attempt fails or no path was recognized, the input text will
    be returned.

    :return: Tuple of a list of the plain input texts or contents of passed pdf file(s) as text and the filenames
    """
    input_buffer = ""

    print("Paste text or insert path to pdf ('*' may be used in filename). Finish input with empty line.")
    input_read: str = input("> ")
    input_buffer += input_read

    while input_read != "" or input_buffer.strip() == "":
        if input_read == "quit":
            print("\nExit keyword detected. Terminating.")
            out_file.close()
            sys.exit(0)

        input_read = input("> ")
        input_buffer += input_read + " "

    input_buffer = input_buffer.strip()

    if input_buffer.split(".")[-1].lower() == "pdf":
        print("\nInterpreting input as path to pdf file(s).")
        try:
            # Get all files matching input with extension '.pdf' or '.PDF' without duplicates
            file_list = glob(input_buffer)
            file_list_upper = glob(".".join(input_buffer.split(".")[:-1]) + ".PDF")
            file_list_upper = list(filter(lambda x: ".".join(x.split(".")[:-1]) + ".pdf" not in file_list,
                                          file_list_upper))

            file_list = list(set(file_list + file_list_upper))

            if len(file_list) == 0:
                raise FileNotFoundError("No files found.")

            file_list = sorted(file_list, key=str.casefold)

            return GMSExtract.read_pdf_multiple(file_list), file_list
        except FileNotFoundError:
            print("File could not be found. Interpreting input as plain text.")
    return [input_buffer], [""]


if __name__ == '__main__':
    # Clear contents of output file
    out_file = open("out.txt", "w+")
    out_file.close()

    try:
        while True:
            text_inputs, input_files = get_input()

            # Only show file name without path in output
            input_files = list(map(lambda file: file.replace("/", "\\").split("\\")[-1], input_files))

            out_file = open("out.txt", "ab+")

            print("\n" + "#" * 150 + "\n")

            out_string = GMSExtract.string_excel_all(*GMSExtract.process_all(text_inputs), input_files)
            print(out_string)

            print("\n" + "#" * 150 + "\n")

            out_file.close()
    except Exception as error:
        print("Unhandled Exception:", error)
        out_file.close()
