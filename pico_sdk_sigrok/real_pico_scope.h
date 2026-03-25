/**
 * @file real_pico_scope.h
 * @brief Brief description of the file.
 *
 * Detailed description of the file.
 *
 * @author Berran Remzi
 * @date 2025-03-23
 */

#ifndef REAL_PICO_SCOPE_H
#define REAL_PICO_SCOPE_H
#include <stdint.h>

// Input divider ratio at ADC input: Vadc = Vin * (Rlow / (Rhigh + Rlow))
// With 866k (high) and 133k (low), ratio is 133/999 ~= 0.13313.
#define REAL_PICO_SCOPE_DIV_NUM 133U
#define REAL_PICO_SCOPE_DIV_DEN 999U

void real_pico_scope_init(void);
void real_pico_scope(void);
uint8_t real_pico_scope_get_gain(void);

#endif /* REAL_PICO_SCOPE_H */
