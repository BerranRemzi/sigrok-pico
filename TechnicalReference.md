# Technical Reference

Build instructions and technical details for sigrok-pico developers.

For branch overview and user-first navigation, see [README.md](README.md).

---

## Table of Contents

1. [Building the Firmware](#building-the-firmware)
2. [Building libsigrok](#building-libsigrok)
3. [Architecture Overview](#architecture-overview)
4. [Debug Interface](#debug-interface)

---

## Building the Firmware

### Prerequisites

Complete the Raspberry Pi PICO C SDK "getting-started-with-pico" guide first.

### Build Steps

```bash
# 1. Clone the repository
git clone https://github.com/pico-coder/sigrok-pico.git
cd sigrok-pico

# 2. Enter firmware directory
cd pico_sdk_sigrok

# 3. Set SDK path
export PICO_SDK_PATH=<path-to-pico-sdk>

# 4. Build
mkdir build && cd build
cmake ..
make
```

### Output

The build produces `pico_sdk_sigrok.uf2` in `pico_sdk_sigrok/build`. Flash this file to your RP2040 Zero using the standard UF2 bootloader method.

---

## Building libsigrok

### Current Status

The raspberrypi_pico driver is merged into mainline libsigrok as of September 2023.

**Recommended**: Install from [sigrok.org/downloads](https://sigrok.org/wiki/Downloads) rather than building from source.

### Build Notes

If building from source:

1. Follow the official libsigrok build instructions
2. Resolve library dependencies for your platform
3. Build sigrok-cli first to verify USB/serial libraries

> **Warning**: Resolving library dependencies is the most challenging part of building PulseView. This is intrinsic to the project, not specific to the PICO driver.

### Resources

- [libsigrok repository](https://github.com/sigrokproject/libsigrok)
- [sigrok download page](https://sigrok.org/wiki/Downloads)

---

## Architecture Overview

### Sample Rate System

The PIO and ADC share a common sample rate because:
- libsigrok supports only one rate per device
- Simplifies DMA implementation

### Clock Sources

| Subsystem | Clock | Max Rate |
|-----------|-------|----------|
| PIO | sysclk (120 MHz) | 120 Msps |
| ADC | USB clock (48 MHz) | 500 Ksps |

### DMA Implementation

The DMA engine reads from PIO FIFO and writes to memory. For 8+ digital channels at high rates, the DMA requires read-modify-write operations, limiting reliable operation to ≤60 Msps.

---

## Debug Interface

### UART0 Output

The hardware UART0 provides debug output on TX pin.

| Version | Baud Rate |
|---------|-----------|
| rev1 | 115200 |
| rev2+ | 921600 |

### Usage

Connect a USB-UART adapter to UART0 TX and open a terminal at the appropriate baud rate. Debug messages include:
- Configuration changes
- Sample rate settings
- Error conditions

> **Note**: Debug output is optional for normal operation. The sigrok driver reports most user errors directly.

### For Bug Reports

When filing bugs, include UART0 debug output if possible. This helps diagnose issues that may not appear in the sigrok driver output.

---

## Related Documentation

- [Analyzer Guide](AnalyzerGuide.md) - User operations guide
- [Serial Protocol](SerialProtocol.md) - Wire protocol specification
- [Getting Started](GettingStarted.md) - Initial setup guide
