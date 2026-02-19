package main

import (
	"errors"
	"fmt"
	"net/http"
	"os"

	"github.com/witanlabs/witan-cli/client"
	"github.com/witanlabs/witan-cli/cmd"
)

func main() {
	if err := cmd.Execute(); err != nil {
		var exitErr *cmd.ExitError
		if errors.As(err, &exitErr) {
			os.Exit(exitErr.Code)
		}

		var apiErr *client.APIError
		if errors.As(err, &apiErr) && apiErr.StatusCode == http.StatusTooManyRequests {
			fmt.Fprintln(os.Stdout, apiErr.Error())
			os.Exit(1)
		}

		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
