# Analyzer Guide

Complete reference for using the sigrok-pico logic analyzer and oscilloscope.

For branch context and hardware scope, see [README.md](README.md).

> **Tip**: Run sigrok-cli or PulseView at debug level 2 (`-l 2`) to see configuration issues.

---

## Table of Contents

1. [Channels](#channels)
2. [Trigger Modes](#trigger-modes)
3. [Storage Modes](#storage-modes)
4. [Sample Rates](#sample-rates)
5. [Best Practices](#best-practices)

---

## Channels

### Digital Channels (7)

- Channel range: D2-D8
- Pin source: board digital pins
- Rule: channels must be enabled contiguously starting from D2

**Enable Rule**: Channels must be enabled sequentially from D2 toward D8. You cannot skip channels.

### Analog Channels (2)

- A0 uses ADC0 on pin 31
- A1 uses ADC1 on pin 32

**Accuracy**: 7-bit effective (128 divisions, ~20mV resolution)

**Why 7-bit?**
1. RP2040 ADC ENOB is ~8 bits despite 12-bit output
2. 7-bit encoding avoids problematic ASCII characters in serial transfer
3. 128 divisions provide sufficient resolution for most applications

### Disabling Channels

Always disable unused channels to:
- Reduce serial transfer overhead
- Allocate more trace storage to enabled signals

---

## Trigger Modes

### Software Trigger (Default)

Any enabled digital pin can serve as a trigger source. Analog channels are captured synchronously with digital triggers.

**Supported Types**: level, rising, falling, changing

**Pre-capture Ratio**: 0-100% configurable

> **Note**: Software triggering adds host-side processing overhead, which may limit maximum streaming sample rates on slower systems.

### Always Trigger

When no trigger is specified, the device immediately captures a fixed-length trace.

### Hardware Trigger (Removed)

HW triggering via PIO was removed in rev2 due to:
- False triggers when conditions weren't present
- No clear UI method to specify HW vs SW triggering
- Processing overhead reduced streaming rate below SW trigger performance

The trigger conditions are still sent to the device, but HW triggering provides no benefit over host-side SW triggering.

---

## Storage Modes

### Fixed Depth Mode

**Enabled when**: No SW trigger AND sample count fits in device storage

**Advantages**:
- Guaranteed capture completion
- No data loss regardless of USB transfer rate

**Use case**: Simple captures with known sample counts

### Continuous Streaming Mode

**Enabled when**: SW triggering active OR sample count exceeds internal storage

**Characteristics**:
- Data streams to host during capture
- May overflow if USB bandwidth insufficient
- Device detects overflow and sends abort code

**Trade-off**: Larger capture depth vs. risk of data loss

**Overflow Handling**: Aborts reduce sample count but prevent corrupted data from being sent.

---

## Sample Rates

### Quick Reference Table

| Digital | Analog | Samples | Max Rate | Limiting Factor |
|---------|--------|---------|----------|-----------------|
| 1-4 | 0 | ≤400K | 120 Msps | PIO |
| 1-4 | 0 | >400K | 500 Ksps+RLE | USB w/ RLE |
| 5-7 | 0 | ≤200K | 120 Msps | PIO |
| 5-7 | 0 | >200K | 500 Ksps+RLE | USB w/ RLE |
| 8-14 | 0 | ≤100K | 120 Msps | PIO |
| 8-14 | 0 | >100K | 250 Ksps+RLE | USB w/ RLE |
| 15-21 | 0 | ≤50K | 120 Msps | PIO |
| 15-21 | 0 | >50K | 167 Ksps+RLE | USB w/ RLE |
| 0 | 1 | ≤200K | 500 Ksps | ADC |
| 0 | 1 | >200K | 500 Ksps | USB & ADC |
| 0 | 2 | ≤100K | 250 Ksps | ADC |
| 0 | 2 | >100K | 250 Ksps | USB & ADC |
| 0 | 3 | ≤67K | 160 Ksps | ADC |
| 0 | 3 | >67K | 160 Ksps | USB & ADC |
| 1-7 | 1 | ≤100K | 500 Ksps | ADC |
| 1-7 | 1 | >100K | 250 Ksps | USB |
| 1-7 | 2 | ≤67K | 250 Ksps | ADC |
| 1-7 | 2 | >67K | 160 Ksps | ADC & USB |
| 1-7 | 3 | ≤50K | 160 Ksps | ADC |
| 1-7 | 3 | >50K | 125 Ksps | USB & ADC |
| 8-14 | 1 | ≤67K | 500 Ksps | ADC |
| 8-14 | 1 | >67K | 160 Ksps | USB |
| 8-14 | 2 | ≤50K | 250 Ksps | ADC |
| 8-14 | 2 | >50K | 125 Ksps | USB |
| 8-14 | 3 | ≤40K | 160 Ksps | ADC |
| 8-14 | 3 | >40K | 100 Ksps | USB |

### Limiting Factors

| Factor | Limit | Description |
|--------|-------|-------------|
| **PIO** | 120 MHz | Maximum system clock for programmable I/O |
| **ADC** | 500 Ksps | ADC converter rate, shared across channels |
| **USB** | 400-800 KB/sec | CDC serial transfer rate (host-dependent) |
| **USB w/ RLE** | Variable | RLE compression reduces wire bytes |

### RLE Optimization

Run Length Encoding reduces bandwidth for signals with low activity factors:

```
Effective rate = listed max × (1 / activity_factor)
```

**Example**: 25% activity factor with 1-4 digital channels can support ~2 Msps.

The RLE algorithm for 1-4 channels is more efficient than for 5-21 channels.

### Hard Limits

**Common Sample Rate**: PIO and ADC share one sample rate (libsigrok limitation).

**Granularity**:
- ADC: Integer divisor of 48 MHz USB clock (worst case: 5 KHz granularity)
- PIO: Full fractional divisor of 120 MHz sysclk

PulseView provides only frequencies yielding integer divisors to prevent clock drift.

**Range**:
- Minimum: 5 KHz (16-bit divisor limit)
- Maximum: 120 Msps (digital only)
- Recommended: ≤60 Msps for 8+ channels (DMA read-modify-write requirement)

### Soft Limits

Soft limits depend on:
- USB link bandwidth
- Device processing capacity
- Host processing speed (especially with SW triggers)

**Avoid soft limits by**: Using Fixed Depth mode with appropriate sample counts.

---

## Best Practices

### Input Protection

Use ≥1 kΩ resistors inline between signal sources and PICO inputs to prevent damage from:
- Voltages outside 0V-3.3V range
- Excessive current from low-impedance sources

### Rate Selection

1. Start with lower sample rates and increase as needed
2. For digital-only captures with sparse signals, leverage RLE
3. Check debug output at `-l 2` for rate limit warnings

### Debug UART

UART0 TX outputs debug information:
- rev1: 115200 bps
- rev2+: 921600 bps

Not required for normal operation, but useful for bug reports.

---

## Related Documentation

- [Getting Started](GettingStarted.md) - Initial setup guide
- [Serial Protocol](SerialProtocol.md) - Wire protocol specification
- [Technical Reference](TechnicalReference.md) - Build instructions and architecture
