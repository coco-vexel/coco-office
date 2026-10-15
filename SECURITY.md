# Security policy

Do not put unpatched vulnerabilities, private documents or credentials in public
issues. Use GitHub's **Security → Report a vulnerability** to open a private
security advisory with the maintainers. Do not open a public issue or pull
request containing exploit details.

Include the affected version, a minimal synthetic reproducer, impact and relevant
runtime details. Avoid real documents and remove credentials and personal data.

The Word and Excel `executeCode` APIs intentionally execute trusted host code and
are not security sandboxes. Applications must not pass document content,
clipboard data, network responses, raw user input, or unreviewed generated code
as executable source. Pass untrusted values through the structured `args`
parameter and prefer the typed live-object APIs.

## Supported versions

Security fixes target the latest published release and the current `main`
branch. Older releases may be assessed case by case and are not guaranteed to
receive backports.

## What to expect

Maintainers will acknowledge a report when it is reviewed, keep discussion in
the private advisory, and coordinate disclosure after a fix or mitigation is
available. Response and remediation times depend on severity and maintainer
availability; this project does not currently publish a response-time SLA.
