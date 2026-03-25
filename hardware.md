# RealPicoScope Hardware Deep Dive

This document contains detailed hardware notes for this branch.
Start from [README.md](README.md) for primary setup and document navigation.

## Video Link

[https://www.youtube.com/](https://www.youtube.com/)

## Project Introduction

This project transforms RP2040 Zero into a low-cost logic analyzer and oscilloscope integrated with sigrok/PulseView.
The design adds input protection, programmable analog gain, and gain status LEDs for practical bench use.

## Project Function

- Logic analyzer: 7 digital channels for protocol analysis (I2C, SPI, UART, and related buses)
- Oscilloscope: 2 analog channels with programmable gain (1, 2, 4, 5, 8, 10, 16, 32)
- Mixed-signal capture: simultaneous digital and analog sampling
- Test output: onboard 1 kHz square wave for verification and probe checks
- LED indicators: visual feedback for active gain setting
- Host tools: PulseView GUI and sigrok-cli support on Windows, Linux, and macOS

## Project Parameters

- MCU: RP2040 Zero (Waveshare)
- Analog front-end: 2x MCP6S21 programmable gain amplifiers
- ADC input divider: 866k/133k network for high impedance and expanded measurable range
- Digital channels: 7 protected inputs (D2-D8)
- Digital protection: 10k series resistors with clamp diodes
- Sample rates: up to 120 Msps digital (PIO), up to 500 Ksps analog (ADC)
- Gain indication: LED display tracks current gain state
- Test signal: 1 kHz square wave output
- Software compatibility: sigrok/PulseView ecosystem

## Hardware Principle

The hardware is organized in these functional blocks:

- Power supply: USB-powered through RP2040 Zero regulator path
- Analog input stage: divider + MCP6S21 PGA per channel, software-selectable gain
- Digital protection: series resistor and clamp path on each digital input
- Test signal generator: dedicated GPIO PWM output at 1 kHz
- Main controller: RP2040 firmware handles sampling, gain control, and USB CDC link

## Software Links

Firmware and host-side code:

- Local branch firmware source: [pico_sdk_sigrok](pico_sdk_sigrok)
- Protocol and usage docs: [SerialProtocol.md](SerialProtocol.md), [AnalyzerGuide.md](AnalyzerGuide.md)

## Safety Notes

- Respect voltage domains: digital inputs are 0-3.3V logic domain
- Validate analog input range against your exact divider and gain configuration
- Use the 1 kHz test output to confirm wiring before critical captures

## Assembly Sequence

1. Solder the two MCP6S21 PGA ICs first.
2. Populate passive components (resistors, diodes, capacitors).
3. Add gain-indicator LEDs and their current-limiting resistors.
4. Add input protection parts and input connectors.
5. Mount and solder the RP2040 Zero module.
6. Install final board connector/header.
7. Flash firmware and verify operation using the 1 kHz test output.

## Finished Product

Current media assets:

- Assembly photos: [hw/image](hw/image)
- 3D board render: [hw/3d/pcba-3d.png](hw/3d/pcba-3d.png)
- Schematic image: [hw/schematic/schematic.png](hw/schematic/schematic.png)
