# Getting Started

Quick setup guide for the RealPicoScope branch of sigrok-pico.

This guide assumes the custom RP2040 Zero hardware used in this branch. For branch overview and document map, see [README.md](README.md).

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation](#installation)
3. [First Capture](#first-capture)
4. [Troubleshooting](#troubleshooting)

---

## Prerequisites

### Hardware

- RealPicoScope board (Waveshare RP2040 Zero based)
- USB cable (data-capable)
- Test leads/probes for digital and analog channels

> **Warning**: Respect branch hardware limits. Digital inputs are 0-3.3V domain. Analog limits depend on the divider/gain front-end used on your board build.

### Software

Install PulseView or sigrok-cli from [sigrok.org/downloads](https://sigrok.org/wiki/Downloads).

---

## Installation

### Step 1: Flash the Firmware

1. Hold the BOOTSEL button while connecting the PICO via USB
2. Copy `pico_sdk_sigrok.uf2` to the RPI-RP2 drive
3. The PICO will reboot automatically

### Step 2: Verify Connection

```bash
# List available serial ports
sigrok-cli --list-serial

# Scan for the device (replace /dev/ttyACM0 with your port)
sigrok-cli -l 2 -d raspberrypi-pico:conn=/dev/ttyACM0:serialcomm=115200/flow=0 --scan
```

**Note**: The baud rate parameter is ignored for CDC serial devices.

---

## First Capture

### Using sigrok-cli

```bash
# Basic 4-channel digital capture at 10 KHz (D2-D5)
sigrok-cli -l 2 \
  -d raspberrypi-pico:conn=/dev/ttyACM0:serialcomm=115200/flow=0 \
  --config samplerate=10000 \
  --channels D2,D3,D4,D5 \
  --samples 1000
```

### Using PulseView

1. Launch PulseView
2. Click "Connect to Device"
3. Select "raspberrypi_pico" driver
4. Set the serial port (e.g., `/dev/ttyACM0` or `COM3`)
5. Configure sample rate and channels
6. Click "Run"

> **Tip**: This branch uses 7 digital channels (D2-D8) and 2 analog channels (A0-A1).

---

## Troubleshooting

### Windows Serial Port Issues

Windows serial port access can be problematic. Try these steps in order:

1. **Close conflicting applications** - Windows doesn't allow multiple apps to access the same port

1. **Reboot after installation** - Restart after installing PulseView

1. **Check USB driver** - Zadig may be required to map the USB device (not always needed)

1. **Cycle the connection** - Unplug/replug the PICO and restart PulseView

1. **Test with a terminal** - Open a serial terminal (TeraTerm, PuTTY), send `*` then `i` to verify device response:

  ```text
   > *
   > i
   SRPICO,A03D21,00
   ```

1. **Repeat steps** - The issue often resolves after several attempts

### Common Problems

- Device not found: check USB cable is data-capable and try a different USB port.
- No serial ports listed: ensure firmware is flashed and BOOTSEL is not held.
- Connection timeout: close other serial applications, reboot, then retry terminal test.
- Sample-rate errors: see [AnalyzerGuide.md](AnalyzerGuide.md) for capture limits.

### Debug Output

For detailed diagnostics, run with debug level 2:

```bash
sigrok-cli -l 2 ...
```

---

## Next Steps

- Read [AnalyzerGuide.md](AnalyzerGuide.md) for channel configuration, trigger modes, and sample rate details
- See [TechnicalReference.md](TechnicalReference.md) for build instructions
- Refer to [SerialProtocol.md](SerialProtocol.md) for protocol documentation
