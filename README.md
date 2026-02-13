# sigrok-pico

Use a Raspberry Pi PICO (RP2040) as a logic analyzer and oscilloscope with sigrok.

---

## Project Status

**Merged to mainline sigrok** (September 2023)

Install from [sigrok.org/downloads](https://sigrok.org/wiki/Downloads) for the recommended experience.

- Pull request: [#181](https://github.com/sigrokproject/libsigrok/pull/181)

---

## Overview

This project implements a sigrok driver for the Raspberry Pi PICO RP2040 using the PICO SDK CDC serial library. It enables the PICO to function as:

- **21-channel logic analyzer** (digital pins D2-D22)
- **3-channel oscilloscope** (analog pins A0-A2)
- **Mixed-signal analyzer** (combined digital + analog)

---

## Documentation

| Document | Purpose |
|----------|---------|
| [Getting Started](GettingStarted.md) | Initial setup and first capture |
| [Analyzer Guide](AnalyzerGuide.md) | Complete operations reference |
| [Technical Reference](TechnicalReference.md) | Build instructions and architecture |
| [Serial Protocol](SerialProtocol.md) | Wire protocol specification |

---

## Directory Structure

```
sigrok-pico/
├── pico_pgen/          # Digital function generator for testing
├── pico_sdk_sigrok/    # RP2040 firmware source code
└── pulseview/          # Windows installer (unofficial)
```

---

## Quick Start

1. Flash `pico_sdk_sigrok.uf2` to your PICO
2. Install PulseView or sigrok-cli from sigrok.org
3. See [GettingStarted.md](GettingStarted.md) for detailed instructions

---

## Building from Source

Building is not recommended for most users. If required, see:

- [TechnicalReference.md](TechnicalReference.md) for firmware build instructions
- [pulseview/Readme.md](pulseview/Readme.md) for Windows installer notes

---

## License

See [LICENSE](LICENSE) for details.
