# Running Witan in Claude Cowork (and other sandboxes)

Witan is a single binary, but it isn't preinstalled in the Cowork Linux sandbox. Set it up once per environment.

## Install

```bash
pip install witan --break-system-packages
```

If the install can't reach the network, the sandbox likely has egress disabled. Ask the user to open **Settings → Capabilities → Allow network egress**, enable **Package managers**, and add **`*.witanlabs.com`** to **Additional allowed domains**. Capability changes apply only to _new_ conversations — they'll need to start a fresh one before retrying.

(`npx witan ...` also works anywhere Node is available, if you'd rather not install.)

## Put it on PATH

A user-level pip install drops the `witan` executable in `~/.local/bin`, which isn't on the sandbox PATH by default, and each bash call starts from a clean environment. **Begin every bash call that uses `witan` with the PATH export** — including the first sanity check:

```bash
export PATH="$HOME/.local/bin:$PATH"
witan --version
```

## Network at runtime

`witan` commands call `*.witanlabs.com`. A `403` with `X-Proxy-Error: blocked-by-allowlist` means the domain isn't allowlisted yet — surface the Settings steps above to the user rather than retrying the command.

## Credentials

Witan runs unauthenticated for local file operations. For organization-backed requests, an API key is read from `WITAN_API_KEY` (or `witan auth login`). If a command reports an auth error, the key is missing or expired — ask the user to set it.
