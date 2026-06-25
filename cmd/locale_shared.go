package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
)

// resolveLocale resolves the locale from flag, environment variables, or returns empty string.
// It checks (in order): explicit flag, WITAN_LOCALE, LC_ALL, LC_MESSAGES, LANG.
// If validateFlag is true and the flag is set, an error is returned for invalid flag values.
// If validateWitanLocale is true, an error is returned for invalid WITAN_LOCALE values.
func resolveLocale(cmd *cobra.Command, flagName, flagValue string, validateFlag, validateWitanLocale bool) (string, error) {
	if cmd.Flags().Changed(flagName) {
		if validateFlag {
			locale, ok := normalizeLocale(flagValue)
			if !ok {
				return "", fmt.Errorf("invalid --%s %q", flagName, flagValue)
			}
			return locale, nil
		}
		return flagValue, nil
	}

	if raw, ok := os.LookupEnv("WITAN_LOCALE"); ok && strings.TrimSpace(raw) != "" {
		if validateWitanLocale {
			locale, valid := normalizeLocale(raw)
			if !valid {
				return "", fmt.Errorf("invalid WITAN_LOCALE %q", raw)
			}
			return locale, nil
		}
		locale, _ := normalizeLocale(raw)
		return locale, nil
	}

	if raw, ok := os.LookupEnv("LC_ALL"); ok && strings.TrimSpace(raw) != "" {
		locale, _ := normalizeLocale(raw)
		return locale, nil
	}

	for _, key := range []string{"LC_MESSAGES", "LANG"} {
		raw, ok := os.LookupEnv(key)
		if !ok || strings.TrimSpace(raw) == "" {
			continue
		}
		if locale, valid := normalizeLocale(raw); valid {
			return locale, nil
		}
	}

	return "", nil
}

// normalizeLocale normalizes a locale string to BCP 47 format.
// Returns (normalized, true) if valid, ("", false) if invalid.
func normalizeLocale(raw string) (string, bool) {
	value := strings.TrimSpace(raw)
	if value == "" {
		return "", true
	}

	if idx := strings.IndexByte(value, '.'); idx >= 0 {
		value = value[:idx]
	}
	if idx := strings.IndexByte(value, '@'); idx >= 0 {
		value = value[:idx]
	}

	value = strings.TrimSpace(strings.ReplaceAll(value, "_", "-"))
	if value == "" {
		return "", true
	}

	upper := strings.ToUpper(value)
	if upper == "C" || upper == "POSIX" || value == "*" {
		return "", false
	}

	parts := strings.Split(value, "-")
	for i, part := range parts {
		if part == "" || !isLocaleToken(part) {
			return "", false
		}
		switch {
		case i == 0:
			parts[i] = strings.ToLower(part)
		case len(part) == 4 && isAlpha(part):
			parts[i] = strings.ToUpper(part[:1]) + strings.ToLower(part[1:])
		case len(part) == 2 && isAlpha(part):
			parts[i] = strings.ToUpper(part)
		case len(part) == 3 && isNumeric(part):
			parts[i] = part
		default:
			parts[i] = strings.ToLower(part)
		}
	}

	return strings.Join(parts, "-"), true
}

func isLocaleToken(s string) bool {
	for _, r := range s {
		if (r < 'a' || r > 'z') && (r < 'A' || r > 'Z') && (r < '0' || r > '9') {
			return false
		}
	}
	return true
}

func isAlpha(s string) bool {
	for _, r := range s {
		if (r < 'a' || r > 'z') && (r < 'A' || r > 'Z') {
			return false
		}
	}
	return true
}

func isNumeric(s string) bool {
	for _, r := range s {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}
