.PHONY: validate version check-tag install-runtime install-plugin preflight

CONTAINER ?= euro-office

VERSION = $(shell python3 -c 'import json; print(json.load(open("plugin/cloudcix-guiden/config.json"))["version"])')

validate:
	grep -Eq '^FROM ghcr\.io/euro-office/documentserver:v[0-9.]+@sha256:[0-9a-f]{64}$$' Dockerfile
	python3 -m json.tool plugin/cloudcix-guiden/config.json >/dev/null
	@! grep -ho '?v=[0-9.]*"' plugin/cloudcix-guiden/*.html | grep -v '?v=$(VERSION)"' || { echo 'The ?v= in the plugin HTML must be the version in config.json, $(VERSION), so browsers fetch the new files' >&2; exit 1; }
	@grep -q '^## $(VERSION) ' CHANGELOG.md || { echo 'CHANGELOG.md has no section for $(VERSION), the version in config.json' >&2; exit 1; }
	bash -n scripts/preflight-runtime.sh
	bash -n scripts/install-runtime.sh
	bash -n scripts/install-plugin.sh

# The plugin's version, from config.json
version:
	@echo $(VERSION)

# A release tag must be v<version in config.json>: make check-tag TAG=v2.5.1
check-tag:
	@test "$(TAG)" = "v$(VERSION)" || { echo 'Tag $(TAG) does not match config.json version $(VERSION)' >&2; exit 1; }

preflight:
	./scripts/preflight-runtime.sh $(CONTAINER)

install-runtime:
	./scripts/install-runtime.sh $(CONTAINER)

install-plugin:
	./scripts/install-plugin.sh $(CONTAINER)
