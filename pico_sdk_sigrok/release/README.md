# Firmware Release Notes

UF2 files for the sigrok-pico project.

> **Note**: All UF2 files are released here rather than in `/build`.

---

## Stable Releases (2023)

| File | Description |
|------|-------------|
| `pico_original.uf2` | Original release (~2023). Supports RP2040/PICO only, 21 digital + 3 analog + D0/D1 as UART |
| `pico_original_serial.uf2` | PR #63 build. Same as above with TUD serial config override |

---

## Major 2025 Refactor

The 2025 refactor includes:
- Redone DMA programming
- RP2350 (PICO 2) support
- 26 and 32 pin digital input modes

> **Note**: Current PulseView releases always start digital pin names at "D2". Even though dig26/dig32 enable D0/D1 as digital inputs, they will show up starting at D2.

### RP2040/PICO Variants

| File | Description |
|------|-------------|
| `pico_baseline.uf2` | Baseline: 21 digital + 3 analog |
| `pico_dig26.uf2` | 26 digital (includes D0/D1) |
| `pico_dig32.uf2` | 32 digital (includes D0/D1) |

### RP2350/PICO2 Variants

| File | Description |
|------|-------------|
| `pico2_baseline.uf2` | Baseline: 21 digital + 3 analog |
| `pico2_dig26.uf2` | 26 digital (includes D0/D1) |
| `pico2_dig32.uf2` | 32 digital (includes D0/D1) |

---

## Recommendation

For new projects, use the 2025 refactor UF2 files for RP2350 support or 26/32 digital pin support.
