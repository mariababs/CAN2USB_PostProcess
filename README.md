# CAN2USB_PostProcess
Post-processing script for data captured by innomaker CAN2USB converter

- Meant to be run on the .xls file that is can be output/saved from the innomaker CAN application.
- The Innomaker CAN application should already be installed on the proplab computer but for some reason needs admin sign-in to run
- To set up the Innomaker CAN application to receive CAN data, make sure the third-part can-to-usb device is plugged in with the CANH, CANL, and GND connected to it.
- You will need to selected the device (It should be the only device that shows up in the drop down), set the baud rate to 125k and then 'open' the device and you should see datarting the come in
- To "record" this data, simply hit the 'clear' button. When you export, it will export every packet since the last time you hit clear or opened the application.
