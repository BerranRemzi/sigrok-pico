# Serial Protocol Specification

Wire protocol for communication between the sigrok driver and the PICO device.

---

## Table of Contents

1. [Overview](#overview)
2. [Control Commands](#control-commands)
3. [Data Transfer Protocols](#data-transfer-protocols)
4. [Device-to-Host Commands](#device-to-host-commands)

---

## Overview

The serial protocol consists of four transfer flows:

| Flow | Purpose |
|------|---------|
| Configuration | Set sample rates, channels, triggers |
| General Data | Analog or >4 digital channels |
| Optimized Data | ≤4 digital channels with RLE |
| Final Count | Byte count at transfer end |

---

## Control Commands

### Immediate Commands (No Response)

| Command | Description |
|---------|-------------|
| `*` | **Reset** - Terminate sampling, clear state, return to idle. Does not affect USB CDC link. Sent by driver during device scan and at acquisition start. |
| `+` | **Abort** - Host-forced abort (PulseView stop button). Terminates data send with minimal state restoration. |

### Commands Requiring Response

Send command followed by `\n` or `\r`. Device returns a string response or times out on error.

| Command | Response Format | Description |
|---------|-----------------|-------------|
| `i` | `SRPICO,AxxDyy,00` | **Identify** - `Axx` = analog channel count, `Dyy` = digital channel count. Example: `SRPICO,A03D21,00` |
| `aX` | `aaaaxbbbbb` | **Analog Scale** - Returns scale/offset for channel X. `aaaa` = scale (µV), `bbbbb` = offset (µV). Combined length ≤18 chars. Supports negative signs. |

### Commands with ACK Response

Device returns `*` on success, nothing on error (driver timeout).

| Command | Format | Description |
|---------|--------|-------------|
| `R` | `R<rate>` | **Sample Rate** - Decimal value, e.g., `R100000` |
| `L` | `L<count>` | **Sample Limit** - Decimal value, e.g., `L5000` |
| `A` | `A<en><ch>` | **Analog Channel** - `en`: 0=disable, 1=enable. `ch`: channel number. Example: `A103` enables A3 |
| `D` | `D<en><ch>` | **Digital Channel** - Same format as analog. Example: `D020` disables D20 |

### Commands Without Response

| Command | Description |
|---------|-------------|
| `F` | **Fixed Sample Mode** - Capture fixed sample set (no SW triggering) |
| `C` | **Continuous Mode** - Stream data for SW trigger processing |

---

## Data Transfer Protocols

### General Protocol

Used when analog channels are enabled OR >4 digital channels are enabled.

**Format**: Samples sent as time-synchronized slices (all channel values for one time point).

**Encoding**: Each byte OR'd with `0x80` to avoid ASCII control characters.

**Digital**: 7 channels per byte, lowest channels first.

**Analog**: 7-bit sample value per channel.

**Example**: 14 digital channels (D2-D15) + 2 analog channels (A0-A1):

```
Slice: 0x8F, 0xA3, 0x91, 0xB6
        │     │     │     └── A1 = 0x36
        │     │     └──────── A0 = 0x11
        │     └────────────── D15:D9 = 0x23
        └──────────────────── D8:D2 = 0x0F
```

### Optimized Protocol (RLE)

Used for ≤4 digital channels with no analog.

**Purpose**: Enable high-speed sampling of narrow-width protocols (I2C, I2S, SPI) by reducing wire bandwidth.

**Encoding**: Each byte contains:
- 4-bit sample value
- RLE count (0-7 samples inline, or 8-640 samples extended)

This protocol allows sample rates exceeding the raw USB transfer capacity when signal activity is sparse.

---

## Device-to-Host Commands

| Command | Description |
|---------|-------------|
| `!` | **Abort Signal** - Device detected capture overflow. Sent periodically until host sends `*` or `+`. Indicates data loss occurred; sample count reduced but remaining data is valid. |

---

## Related Documentation

- [Analyzer Guide](AnalyzerGuide.md) - User operations guide
- [Technical Reference](TechnicalReference.md) - Build instructions and architecture
