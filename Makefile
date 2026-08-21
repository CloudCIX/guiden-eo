.PHONY: validate install-runtime install-plugin preflight

CONTAINER ?= euro-office

validate:
	grep -Eq '^FROM ghcr\.io/euro-office/documentserver@sha256:[0-9a-f]{64}$$' Dockerfile
	python3 -m json.tool plugin/cloudcix-guiden/config.json >/dev/null
	bash -n scripts/preflight-runtime.sh
	bash -n scripts/install-runtime.sh
	bash -n scripts/install-plugin.sh


preflight:
	./scripts/preflight-runtime.sh $(CONTAINER)

install-runtime:
	./scripts/install-runtime.sh $(CONTAINER)

install-plugin:
	./scripts/install-plugin.sh $(CONTAINER)
