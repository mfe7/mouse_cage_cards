# mouse_cage_cards

This repo contains a python script that takes in a cage list from SoftMouse.NET and outputs an excel spreadsheet that contains notecards that can be placed at each cage.
The cards can be further customized to display relevant information, such as contact info.

Here are instructions to generate the mouse cage cards:

## Set up for the first time

#### Get the code from github
In a terminal do the following commands:
```bash
cd ~
git clone https://github.com/mfe7/mouse_cage_cards.git
```

#### Install the required python libraries
Only the very first time you've downloaded the code, run this in the terminal:
```bash
pip install xlsxwriter xlrd
```

#### Set up your contact info 

Create a file called `settings.yaml` in this repo (`~/mouse_cage_cards`) that has the following info:
```yaml
PI_name: pi_last_name
protocol_num: 00000
contact_name: 'Bob Smith'
contact_phone: '(123) 555-1234'
species: Mouse
```

## To update your cage cards

#### Get the Cage List from SoftMouse.NET
* Click "Cages" in the top bar
* Export the Cage List as a spreadsheet
* Re-name the downloaded spreadsheet to `softmousedb.xlsx` and place it in this folder (`~/mouse_cage_cards`)

#### Re-generate the notecards
In a terminal, run:
```bash
cd ~/mouse_cage_cards
python notecard.py
```

* That will print out how many pieces of paper to load into the printer (split by mouseline)
* The python also generates a file called `notecards.xlsx`
* Open and print `notecards.xlsx`

Success!