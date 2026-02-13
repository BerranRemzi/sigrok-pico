/**
 * @file real_pico_scope.c
 * @brief Implementation of analog front-end controls for the Pico-based
 * oscilloscope.
 *
 * This module handles the programmable gain amplifier (PGA) control via SPI,
 * LED display using Charlieplexing, button input for gain cycling, and square
 * wave generation.
 *
 * @author Berran Remzi
 * @date 2025-03-23
 */

#include "real_pico_scope.h"
#include "hardware/gpio.h"
#include "hardware/pwm.h"
#include "pico/stdlib.h"

/* Constants */
#define GAIN_COUNT 8U
#define BUTTON_PIN 13U
#define SQUARE_WAVE_PIN 14U
#define SQUARE_WAVE_FREQ 1000U // 1kHz
#define SPI_CS_PIN 28U
#define SPI_SCK_PIN 29U
#define SPI_MOSI_PIN 15U
#define MCP6S21_CMD_WRITE_GAIN 0x40U // Command to write gain for channel 0

/* Available Gain Values and Codes */
static const uint8_t GAIN_VALUES[GAIN_COUNT] = {1, 2, 4, 5, 8, 10, 16, 32};
static const uint8_t GAIN_CODES[GAIN_COUNT] = {0x00, 0x01, 0x02, 0x03,
                                               0x04, 0x05, 0x06, 0x07};

static uint8_t current_gain_index = 0U;

/**
 * @brief Initializes square wave output on the specified GPIO pin using PWM.
 *
 * Generates a 1kHz square wave at 50% duty cycle.
 */
static void square_wave_init(void) {
  // Set GPIO function to PWM
  gpio_set_function(SQUARE_WAVE_PIN, GPIO_FUNC_PWM);

  // Get PWM slice and channel
  uint slice_num = pwm_gpio_to_slice_num(SQUARE_WAVE_PIN);
  uint channel = pwm_gpio_to_channel(SQUARE_WAVE_PIN);

  // Disable PWM for configuration
  pwm_set_enabled(slice_num, false);

  // Set clock divider for 1MHz base frequency
  pwm_set_clkdiv(slice_num, 120.0f);

  // Set wrap for 1kHz frequency
  pwm_set_wrap(slice_num, 1000U);

  // Set 50% duty cycle
  pwm_set_chan_level(slice_num, channel, 500U);

  // Enable PWM
  pwm_set_enabled(slice_num, true);
}

/**
 * @brief Initializes software SPI GPIOs.
 */
static void spi_init(void) {
  gpio_init(SPI_CS_PIN);
  gpio_init(SPI_SCK_PIN);
  gpio_init(SPI_MOSI_PIN);

  gpio_set_dir(SPI_CS_PIN, GPIO_OUT);
  gpio_set_dir(SPI_SCK_PIN, GPIO_OUT);
  gpio_set_dir(SPI_MOSI_PIN, GPIO_OUT);

  gpio_put(SPI_CS_PIN, true);  // CS idle high
  gpio_put(SPI_SCK_PIN, true); // SCK idle high for mode 3
  gpio_put(SPI_MOSI_PIN, false);
}

/**
 * @brief Generates a clock pulse for SPI mode 3 (CPOL=1, CPHA=1).
 */
static void spi_clock_pulse(void) {
  gpio_put(SPI_SCK_PIN, false); // Falling edge - data sampled
  sleep_us(1);
  gpio_put(SPI_SCK_PIN, true); // Rising edge - data can change
  sleep_us(1);
}

/**
 * @brief Sends one byte over software SPI (MSB first) in mode 3.
 *
 * @param data The byte to send.
 */
static void spi_send_byte(uint8_t data) {
  for (int i = 7; i >= 0; i--) {
    gpio_put(SPI_MOSI_PIN, (data >> i) & 1U);
    sleep_us(1); // Setup time
    spi_clock_pulse();
  }
}

/**
 * @brief Sets the gain of the MCP6S21 PGA.
 *
 * @param gain_code The gain code (0-7) corresponding to gains 1x to 32x.
 */
static void mcp6s21_set_gain(uint8_t gain_code) {
  gpio_put(SPI_CS_PIN, false);           // Select device
  sleep_us(10);                          // Small delay
  spi_send_byte(MCP6S21_CMD_WRITE_GAIN); // Send command
  spi_send_byte(gain_code);              // Send gain value
  sleep_us(10);                          // Small delay
  gpio_put(SPI_CS_PIN, true);            // Deselect device
}

/**
 * @brief Controls the LED display using Charlieplexing.
 *
 * Uses 4 GPIO pins to drive up to 8 LEDs. Each index corresponds to a specific
 * LED configuration.
 *
 * @param index The LED index (0-7) to turn on. Values outside this range turn
 * off all LEDs.
 */
static void led_display(uint8_t index) {
#define LED_PIN_COUNT 4U
#define LED_COUNT 8U

  const uint8_t gpio_pins[LED_PIN_COUNT] = {9U, 10U, 11U, 12U};

  typedef struct {
    uint8_t low_pin_index;
    uint8_t high_pin_index;
  } led_config_t;

  const led_config_t led_configs[LED_COUNT] = {
      {.low_pin_index = 1U, .high_pin_index = 2U},
      {.low_pin_index = 2U, .high_pin_index = 1U},
      {.low_pin_index = 0U, .high_pin_index = 1U},
      {.low_pin_index = 0U, .high_pin_index = 2U},
      {.low_pin_index = 1U, .high_pin_index = 0U},
      {.low_pin_index = 2U, .high_pin_index = 0U},
      {.low_pin_index = 3U, .high_pin_index = 1U},
      {.low_pin_index = 1U, .high_pin_index = 3U},
  };

  // Set all pins to input (high impedance) to turn off LEDs
  for (uint8_t i = 0U; i < LED_PIN_COUNT; i++) {
    gpio_init(gpio_pins[i]);
    gpio_set_pulls(gpio_pins[i], false, false);
    gpio_set_dir(gpio_pins[i], GPIO_IN);
  }

  // If index is valid, configure the specific LED
  if (index < LED_COUNT) {
    gpio_set_dir(gpio_pins[led_configs[index].low_pin_index], GPIO_OUT);
    gpio_set_dir(gpio_pins[led_configs[index].high_pin_index], GPIO_OUT);
    gpio_put(gpio_pins[led_configs[index].low_pin_index], false);
    gpio_put(gpio_pins[led_configs[index].high_pin_index], true);
  }
}

/**
 * @brief Initializes the real pico scope module.
 *
 * Sets up button input, square wave output, SPI, and initial gain/LED state.
 */
void real_pico_scope_init(void) {
  // Initialize button pin
  gpio_init(BUTTON_PIN);
  gpio_set_pulls(BUTTON_PIN, true, false); // Pull-up
  gpio_set_dir(BUTTON_PIN, GPIO_IN);

  // Initialize square wave
  square_wave_init();

  // Initialize SPI
  spi_init();

  // Set initial gain and update LED
  led_display(current_gain_index);
  mcp6s21_set_gain(GAIN_CODES[current_gain_index]);
}

/**
 * @brief Main loop function for the real pico scope module.
 *
 * Handles button presses to cycle through gain settings with simple debouncing.
 */
void real_pico_scope(void) {
  static bool prev_button_state = true;
  bool current_button_state = gpio_get(BUTTON_PIN);

  // Detect button press (falling edge) with debouncing
  if ((current_button_state == false) && (prev_button_state == true)) {
    // Cycle gain index
    if (current_gain_index < (GAIN_COUNT - 1U)) {
      current_gain_index++;
    } else {
      current_gain_index = 0U;
    }

    // Update LED and gain
    led_display(current_gain_index);
    mcp6s21_set_gain(GAIN_CODES[current_gain_index]);

    // Simple debouncing delay
    sleep_ms(100);
  }

  prev_button_state = current_button_state;
}
