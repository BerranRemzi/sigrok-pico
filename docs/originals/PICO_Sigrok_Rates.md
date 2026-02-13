```markdown
# Sample Rate Reference Table

Pick the number of enabled digital and analog channels and the number of samples requested to determine the maximum sample rate and its limiting factor.

---

## Configuration Table

| Digital Channels | Analog Channels | Number of Samples | Max Sample Rate | Limitation |
|------------------|----------------|-------------------|-----------------|------------|
| 1-4  | 0 | <=400K | 120Msps | PIO |
| 1-4  | 0 | >400K  | 500Ksps+RLE | USB w/ RLE |
| 5-7  | 0 | <=200K | 120Msps | PIO |
| 5-7  | 0 | >200K  | 500Ksps+RLE | USB w/ RLE |
| 8-14 | 0 | <=100K | 120Msps | PIO |
| 8-14 | 0 | >100K  | 250Ksps+RLE | USB w/ RLE |
| 15-21| 0 | <=50K  | 120Msps | PIO |
| 15-21| 0 | >=50K  | 167Ksps+RLE | USB w/ RLE |
| 0 | 1 | <=200K | 500Ksps | ADC |
| 0 | 1 | >=200K | 500Ksps | USB&ADC |
| 0 | 2 | <=100K | 250Ksps | ADC |
| 0 | 2 | >=100K | 250Ksps | USB&ADC |
| 0 | 3 | <=67K | 160Ksps | ADC |
| 0 | 3 | >=67K | 160Ksps | USB&ADC |
| 1-7 | 1 | <=100K | 500Ksps | ADC |
| 1-7 | 1 | >=100K | 250Ksps | USB |
| 1-7 | 2 | <=67K | 250Ksps | ADC |
| 1-7 | 2 | >=67K | 160Ksps | ADC&USB |
| 1-7 | 3 | <=50K | 160Ksps | ADC |
| 1-7 | 3 | >=50K | 125Ksps | USB&ADC |
| 8-14 | 1 | <=67K | 500Ksps | ADC |
| 8-14 | 1 | >=67K | 160Ksps | USB |
| 8-14 | 2 | <=50K | 250Ksps | ADC |
| 8-14 | 2 | >=50K | 125Ksps | USB |
| 8-14 | 3 | <=40K | 160Ksps | ADC |
| 8-14 | 3 | >=40K | 100Ksps | USB |

---

## Limitations

- **PIO** – The programmable IO runs at a maximum system clock of **120 MHz**.

- **USB** – The USB transfer rate varies from **400–800 kB/sec** depending on the host.  
  Since USB bandwidth varies per host, the maximum sample rate in USB-limited cases will also vary.

- **USB w/ RLE** – Run Length Encoding (RLE) reduces the effective bytes sent.  
  Roughly, an activity factor of **X%** can support a sample rate of:

      listed max sample rate × (1 / X)

  Example: An activity factor of **25%** could support **2 Msps** with **1–4 digital channels**.  
  The RLE algorithm for **1–4 digital channels** is more wire-efficient than for **5–21 channels**.

- **ADC** – The ADC converter runs at **500 Ksps**, shared across each enabled ADC channel.

```
