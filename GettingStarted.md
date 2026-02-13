Warning: Getting the pico device to work on windows can be an initially frustrating experience, but usually the problem resolves itself.  
Recommended steps to resolve issues:
1) Close out any programs that access serial ports as windows won't allow multiple apps to access a port.
2) Once pulseview is installed, do a reboot.
3) Possibly use zadig to map the USB device (You may need to use Zadig to install the USB driver).
4) Try plugging/unplugging the device and/or opening closing pulseview a few times.
5) Close pulseview, and open a serial device apps (Terraterm/Putty etc).  Send a "*" and then a "i" for identify and you should get a response. For some reason, other apps seem to have more success accessing the device than does the libserial code in pulseview.
6) Go back and repeat any of the steps above.  (Yeah, I know it's frustrating, but eventually it seems to work for most folks).

It is recommended that you read through the AnalyzerDetails.md for specifics on modes of the device, but if you just can't wait:

1) Get some 1Kohm or greater resistors to put inline between the PICO inputs, you don't want to accidentally damage your PICO by applying voltages outside the 0V-3.3V range. 
2) Get the appropriate uf2 from [sigrok-pico/pico_sdk_sigrok/release](./pico_sdk_sigrok/release) file and program it to your PICO devices (examples are online, everywhere).
3) Install Pulseview and skip to step 7, or sigrok-cli - follow descriptions from the sigrok pages for both.
4) Replug the PICO to reset it, and then use sigrok-cli to scan for available serial ports
```
sigrok-cli --list-serial
```
5) Use sigrok-cli to scan for the device based on the serial port you found above. The baudrate doesn't matter because we are CDC serial.
```
sigrok-cli  -l 2 -d raspberrypi-pico:conn=/dev/ttyACM0:serialcomm=115200/flow=0 --scan
```
6) Do a first trace 
```
~/github/sigrok-cli/sigrok-cli  -l 2 -d raspberrypi-pico:conn=/dev/ttyACM0:serialcomm=115200/flow=0 --config samplerate=10000  --channels D2,D3,D4,D5 --samples 1000
```
7) Alternatively, use Pulseview instead of sigrok-cli.
8) Go read AnalyzerDetails.md like you should have to begin with....
