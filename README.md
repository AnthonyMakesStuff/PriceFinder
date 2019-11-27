# PriceFinder
The price calculator program analyses a website's manifest and calculates the estimated
profit to be returned. It currently suports the following websites:
  *  http://britdeals.co.uk/lot/[lot url]
  *  https://www.gemwholesale.co.uk/acatalog/[lot number].htm
  
## Usage
To use the program, simply call it with Python 3 and pass any of the following arguments

### Command Line Arguments
| Command Line Argument | Description           | Example Usage                                                                                                                        |
|-----------------------|-----------------------|--------------------------------------------------------------------------------------------------------------------------------------|
| `-site`, `--s`        | The input website URL | `python3 price_calculator.py -site www.website.com`<br>`python3 price_calculator.py --s www.website.com`                             |
| `-output`, `--o`      | The output XLSX file  | `python3 price_calculator.py -site www.website.com -output data.xlsx`<br>`python3 price_calculator.py --s website.com --o data.xlsx` |
  
### Basic Example
```batch
python3 price_calculator.py -site https://www.britdeals.co.uk/lot/computer-accessories-customer-returns-1pallet-94pcs-uklot-825 -output analysed_data.xlsx
```
