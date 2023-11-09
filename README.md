# Chinese Keyword Extractor

## Author
Eric

## Date
11/8/2023 

## Purpose
This Python script is designed for the extraction and analysis of keywords from specific columns of `.xlsx` files. It focuses on processing text from Chinese sentences, filtering out stop words, and other non-meaningful characters to derive meaningful keywords.

## Assumptions
The script is tailored to process data from the 8th column of provided `.xlsx` files, ignoring other columns for the sake of simplicity in keyword extraction.

## Requirements
To run this script, you need to have Python installed on your system along with the following libraries:
- `jieba`: For Chinese text segmentation.
- `openpyxl`: For reading `.xlsx` files.
- `re`: For regular expression operations.
- `json`: For writing output in JSON format.
- `os`: For operating system dependent functionality.

You can install the necessary Python libraries using `pip`. Open your command line interface and run:

```bash
pip install jieba openpyxl
```

## Installation
Clone the repository or download the script to your local machine. If you are cloning the repository, use:

    git clone https://github.com/goEricBi/HF_word_extractor_chinese.git

Place your `.xlsx` files in the same directory as the script or update the file paths in the `main()` function within the script.

## Usage
To execute the script, navigate to the directory containing the script and your `.xlsx` files and run:

    python Review_Extractor.py

Replace `Review_Extractor.py` with the actual name of your script file.

The script will process the 8th column of the specified `.xlsx` files and output a `result.json` file containing the keywords and their frequency, located in the script's directory.

## Output
After successful execution, you will find a `result.json` file in the same directory as the script. This file includes a sorted list of keywords along with their frequencies, formatted in JSON for easy readability and further analysis.

## Contributions
Contributions to this project are welcome! Feel free to fork the repository, make your changes, and submit a pull request. If you encounter any problems or have suggestions, please open an issue on the repository's issues page.

## License
This project is open-sourced under the MIT License. See the [LICENSE.md](LICENSE.md) file for more details.

## Contact
For any further queries or assistance, please reach out to Eric at xuanjiabi@g.ucla.edu
