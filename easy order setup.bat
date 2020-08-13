echo installing easy order...

cd \
cd users
subst x: %USERPROFILE%
x:
cd desktop
mkdir easy order
cd easy order

https://aka.ms/nugetclidl
nuget.exe install python -Version 3.7.7


python -m pip install xlrd
python -m pip install xlsxwriter
python -m pip install pyinstaller

pyinstaller --hidden-import=xlsxwriter --hidden-import=xlrd -F easy_order.py

mkdir your_orders