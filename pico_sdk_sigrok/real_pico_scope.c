/**
 * @file real_pico_scope.c
 * @brief Brief description of the file.
 *
 * Detailed description of the file.
 *
 * @author Berran Remzi
 * @date 2025-03-23
 */

#include "real_pico_scope.h"
#include "hardware/gpio.h"
#include "pico/stdlib.h"

#ifndef HIGH
#define HIGH 1
#endif

#ifndef LOW
#define LOW 0
#endif
#define GAIN_COUNT 8u
#define BUTTON_PIN 13u
// Define software SPI pins
#define SPI_CS   28  // Chip Select
#define SPI_SCK  29  // Clock
#define SPI_MOSI 15  // MOSI (Data Out)

// MCP6S21 Command Definitions
#define CMD_WRITE_GAIN  0x41  // 0b01000001 (Gain Set Command)

// Available Gain Values
const uint8_t GAIN_VALUES[] = {1, 2, 4, 5, 8, 10, 16, 32};
const uint8_t GAIN_CODES[] = {0b000, 0b001, 0b010, 0b011, 0b100, 0b101, 0b110, 0b111};

static uint8_t gain_index = 0u;

void spi_init(void) {
    /* Initialize GPIOs */
    gpio_init(SPI_CS);
    gpio_init(SPI_SCK);
    gpio_init(SPI_MOSI);

    gpio_set_dir(SPI_CS, GPIO_OUT);
    gpio_set_dir(SPI_SCK, GPIO_OUT);
    gpio_set_dir(SPI_MOSI, GPIO_OUT);

    gpio_put(SPI_CS, HIGH);
    gpio_put(SPI_SCK, LOW);
    gpio_put(SPI_MOSI, LOW);
}
// Function to generate a clock pulse
void spi_clk_pulse() {
    gpio_put(SPI_SCK, HIGH);
    sleep_us(1);
    gpio_put(SPI_SCK, LOW);
    sleep_us(1);
}

// Function to send one byte over software SPI (MSB first)
void spi_send_byte(uint8_t data) {
    for (int i = 7; i >= 0; i--) {
        gpio_put(SPI_MOSI, (data >> i) & 1);
        spi_clk_pulse();
    }
}

// Function to set MCP6S21 gain
void mcp6s21_set_gain(uint8_t gain_code) {
    gpio_put(SPI_CS, LOW);  // Select MCP6S21
    spi_send_byte(CMD_WRITE_GAIN);  // Send command
    spi_send_byte(gain_code);  // Send gain value
    gpio_put(SPI_CS, HIGH);  // Deselect MCP6S21
}
void led(uint8_t index) {
    #define LED_PIN_COUNT 4u
    #define LED_COUNT 8u
    const uint8_t gpio_map[LED_PIN_COUNT] = {
      9u, 10u, 11u, 12u
    };

    typedef struct {
      uint8_t low_index;
      uint8_t high_index;
    } gpio_config_t;

    const gpio_config_t gpio_config[LED_COUNT] = {
        {.low_index = 1, .high_index = 2},
        {.low_index = 2, .high_index = 1},
        {.low_index = 0, .high_index = 1},
        {.low_index = 0, .high_index = 2},
        {.low_index = 1, .high_index = 0},
        {.low_index = 2, .high_index = 0},
        {.low_index = 3, .high_index = 1},
        {.low_index = 1, .high_index = 3},
    };

    /* set all to input (high impedance mode) */
    for (uint8_t i = 0u; i < LED_PIN_COUNT; i++) {
      gpio_init(gpio_map[i]);\
      gpio_set_pulls(gpio_map[i], false, false);
      gpio_set_dir(gpio_map[i], GPIO_IN);
    }
    if (index < LED_COUNT) {
      gpio_set_dir(gpio_map[gpio_config[index].low_index], GPIO_OUT);
      gpio_set_dir(gpio_map[gpio_config[index].high_index], GPIO_OUT);
      gpio_put(gpio_map[gpio_config[index].low_index], LOW);
      gpio_put(gpio_map[gpio_config[index].high_index], HIGH);
    }
}

void real_pico_scope_init(void) {
    /* init button */
    gpio_init(BUTTON_PIN);
    gpio_set_pulls(BUTTON_PIN, true, false);
    gpio_set_dir(BUTTON_PIN, GPIO_IN);

    /* init SPI*/
    //  spi_init();
    led(gain_index);
}

void real_pico_scope(void) {
    static uint8_t prev_button_state = HIGH;
    if ((gpio_get(BUTTON_PIN) == LOW) && (prev_button_state == HIGH)) {
        if (gain_index < GAIN_COUNT-1u) {
            gain_index++;
        } else {
            gain_index = 0u;
        }

        led(gain_index);
        mcp6s21_set_gain(gain_index);
        sleep_ms(100);
    }
    prev_button_state = gpio_get(BUTTON_PIN);
}

