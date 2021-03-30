```
  _____             _  ____                                                              
 | ____|__  __ ___ | ||  _ \   ___   ___  ___   _ __ ___   _ __    ___   ___   ___  _ __ 
 |  _|  \ \/ // _ \| || | | | / _ \ / __|/ _ \ | '_ ` _ \ | '_ \  / _ \ / __| / _ \| '__|
 | |___  >  <|  __/| || |_| ||  __/| (__| (_) || | | | | || |_) || (_) |\__ \|  __/| |   
 |_____|/_/\_\\___||_||____/  \___| \___|\___/ |_| |_| |_|| .__/  \___/ |___/ \___||_|   

? Select a target input file: ~/Documents/exel/static/input.xlsx
? Input prefix for transformed file: -output
? Transform Worksheet - List1 ? Yes
? Input headers row: 1
? Input data separator: ;
? Select a target cell: Emails                          
```

## Description
Exel row decomposer

### Parameters
Some parameters are requested during operation for each worksheet, such parameters will be marked as
`L`. Parameters marked as `*` are required:
 1. Select a target input file`*`:
    - Default: static/input.xlsx
    - Description: The input xlsx file for transformation must be placed in the PWD of the program launch
 2. Input prefix for transformed file:
    - Default: -output
    - Description: The prefix to create the output file
 3. Transform Worksheet - ${Worklist} `L`
    - Default: true
    - Description: Skip worksheet transformation option
 4. Input headers row `L`
    - Default: 1
    - Description: Position of the header row in the worksheet
 5. Input data separator `L`
    - Default: `;`    
    - Description: Decomposition separator followed by row copying
 6. Select a target column`*` `L`
    - Description: The target column of the decomposition is presented as a list of headers row


## Installation

```bash
$ npm install
```

## Running the app

```bash
# development
$ npm run start

# watch mode
$ npm run start:dev

# production mode
$ npm run start:prod
```

## Stay in touch

- Author - [Bohdan Shyrokostup](https://t.me/H6L515)
