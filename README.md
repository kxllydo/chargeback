file structure:
./
    archives/
        year/
            month/
                billing.csv
                chargeback.xlsx

                * billing.csv is raw, untouched download billing report from Azure
                * chargeback.xlsx should be a sheet with ONLY the summary sheet inside
    steps/
        stepone.py
        stepthree.py
        __init__.py (maybe?)
    main.py

**for testing (since git doesnt let us have big files), go to main.py and point archive to where you have the files

main.py -> validates file structure for archives
stepone.py -> goes through billing to aggregate cost by RG
stepthree.py -> goes through summary to make new col, insert RG cost

see.txt is just a file RG discrepancy i wanted to bring up with jason tmr!