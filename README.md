quick solution to search through the `doc`/`docx`/`pdf` files by content.


#### 1. Install Python
from [there](https://www.python.org/downloads/)

#### 2. Install pip
described [here](https://pip.pypa.io/en/stable/installation/)

#### 3. Clone/download this repo

#### 4. Install requirements

```pip install -r requirements.txt```

#### 5. Provide pattern and folder to search

open main.py using any editor and set variable `pattern` (row 26) and `directory` (row 27) to value you want to search and root folder you interested for recursive search by all `doc`, `docx` and `pdf` files in all subfolders.

#### 6. Run the script to start search

```python main.py```


#### 7. To-Do
implement command line arguments parsing to provide `pattern` and `directory`.
