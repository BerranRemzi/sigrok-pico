# sigrok-pico (RealPicoScope Branch)

Hardware-focused sigrok firmware for a custom RP2040 Zero board with protected inputs, programmable analog gain, and direct PulseView/sigrok-cli support.

## What This Branch Is

This branch targets the RealPicoScope hardware variant, not the generic upstream board profile.

- Board target: Waveshare RP2040 Zero
- Digital capture: 7 channels (D2-D8)
- Analog capture: 2 channels (A0-A1)
- Analog front-end: MCP6S21 programmable gain stages (1x to 32x)
- Extra hardware controls: gain button, LED gain indicators, 1 kHz test output

Mainline sigrok support was merged in 2023, but this branch keeps custom hardware integration and calibration behavior specific to the RealPicoScope design.

## Quick Start (General Users)

1. Download or build the firmware UF2 for this branch.
2. Put RP2040 Zero into BOOTSEL mode and copy the UF2 file.
3. Install PulseView/sigrok-cli from [sigrok.org/downloads](https://sigrok.org/wiki/Downloads).
4. Connect to driver `raspberrypi-pico` and start with a low sample rate.
5. Use [GettingStarted.md](GettingStarted.md) for first-capture steps and troubleshooting.

## Hardware Overview

RealPicoScope hardware adds instrumentation features on top of RP2040:

- Analog divider network sized for higher-voltage measurement at the ADC input.
- MCP6S21 gain control (1, 2, 4, 5, 8, 10, 16, 32).
- Protected digital input path for embedded debugging workflows.
- Front-panel workflow helpers: pushbutton gain step and LED gain display.
- Built-in 1 kHz test output for probe checks.

For extended hardware description and assembly notes, see [hardware.md](hardware.md).

## Choose Your Path

- First-time setup and first measurement: [GettingStarted.md](GettingStarted.md)
- Capture modes, trigger behavior, and sample-rate guidance: [AnalyzerGuide.md](AnalyzerGuide.md)
- Build details and firmware architecture: [TechnicalReference.md](TechnicalReference.md)
- Driver wire protocol details: [SerialProtocol.md](SerialProtocol.md)
- Windows PulseView installer notes: [pulseview/Readme.md](pulseview/Readme.md)

## Repository Layout (Logical)

```text
sigrok-pico/
|-- README.md                 # Branch landing page (this file)
|-- GettingStarted.md         # User onboarding and troubleshooting
|-- AnalyzerGuide.md          # Runtime usage limits and best practices
|-- TechnicalReference.md     # Build + firmware architecture
|-- SerialProtocol.md         # Device protocol details
|-- hardware.md               # Detailed hardware/assembly context
|
|-- pico_sdk_sigrok/          # Main firmware source for RealPicoScope
|   |-- real_pico_scope.c     # Hardware controls (gain, LED, button, test wave)
|   |-- real_pico_scope.h     # Divider/gain-related constants and API
|   |-- sr_device.c/.h        # Capture engine and channel configuration
|   `-- build/                # Generated artifacts (uf2/elf/map), not source-of-truth
|
|-- hw/                       # Hardware media (photos, 3D view, schematic)
|   |-- image/
|   |-- 3d/
|   `-- schematic/
|
|-- pico_pgen/                # Optional pulse generator helper project
`-- pulseview/                # Host-side notes/tools (Windows-focused)
```

## Branch-Specific Notes

- Firmware compiles `real_pico_scope.c` into the primary target in [pico_sdk_sigrok/CMakeLists.txt](pico_sdk_sigrok/CMakeLists.txt).
- Pin mapping and channel counts are defined in [pico_sdk_sigrok/sr_device.h](pico_sdk_sigrok/sr_device.h).
- Analog divider constants and gain API are defined in [pico_sdk_sigrok/real_pico_scope.h](pico_sdk_sigrok/real_pico_scope.h).

## License

See [LICENSE](LICENSE).
