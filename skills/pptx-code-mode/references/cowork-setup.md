# Installing Witan in Claude Cowork

Witan is a single-binary CLI for agents; `witan pptx` drives PowerPoint files. To set it up in the Cowork Linux sandbox:

**Install:**

```bash
pip install witan --break-system-packages
```

If `pip install` fails to reach the network, it's likely the sandbox has egress disabled. Ask the user to open **Settings → Capabilities → Allow network egress**, enable **Package managers**, and add **`*.witanlabs.com`** to **Additional allowed domains**. Setting changes only apply to new conversations — they'll need to start a fresh one before retrying.

**Running it:** This is a user install, so the `witan` executable lands in `~/.local/bin`, which isn't on the sandbox PATH by default — and each bash call starts with a fresh environment. Begin every bash call that uses `witan` with the PATH export, including the post-install sanity check:

```bash
export PATH="$HOME/.local/bin:$PATH"
witan --version
```

**Network at runtime:** Witan commands call `*.witanlabs.com`. A `403` with `X-Proxy-Error: blocked-by-allowlist` means it isn't allowlisted yet — surface this to the user (see install step above for the settings path) rather than retrying.
