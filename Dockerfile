# Euro-Office v9.3.7, the version CloudCIX runs, pinned by its multi-arch digest.
# To move to another version, look its digest up with:
#   docker buildx imagetools inspect ghcr.io/euro-office/documentserver:<tag>
FROM ghcr.io/euro-office/documentserver@sha256:39c4cdad8c335d05ea5653926e1fc3a69da006d9c1729936b2b4193140cd2962

LABEL org.opencontainers.image.source="https://github.com/CloudCIX/guiden-eo" \
      org.opencontainers.image.description="Euro-Office Document Server with the CloudCIX Guiden plugin"

COPY runtime/v1/ \
    /var/www/euro-office/documentserver/sdkjs-plugins/v1/

COPY plugin/cloudcix-guiden/ \
    /var/www/euro-office/documentserver/sdkjs-plugins/cloudcix-guiden/
