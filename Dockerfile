# The Euro-Office version CloudCIX runs. The digest (multi-arch) is what is built; the tag is there for
# people and for Dependabot, which opens a pull request when Euro-Office publishes a newer version. To move by hand,
# look the new version's digest up with:
#   docker buildx imagetools inspect ghcr.io/euro-office/documentserver:<tag>
FROM ghcr.io/euro-office/documentserver:v9.3.8@sha256:0f19cb676d48f85b826e199ac9e773ea12130f7d32bfc9a45e86aea320441b8a

LABEL org.opencontainers.image.source="https://github.com/CloudCIX/guiden-eo" \
      org.opencontainers.image.description="Euro-Office Document Server with the CloudCIX Guiden plugin"

COPY runtime/v1/ \
    /var/www/euro-office/documentserver/sdkjs-plugins/v1/

COPY plugin/cloudcix-guiden/ \
    /var/www/euro-office/documentserver/sdkjs-plugins/cloudcix-guiden/

# The editor's web server sends a file's .gz copy when there is one, so remove any the base image had for the files
# above; otherwise browsers would keep getting the image's old copies.
RUN find /var/www/euro-office/documentserver/sdkjs-plugins/v1 \
         /var/www/euro-office/documentserver/sdkjs-plugins/cloudcix-guiden \
         -name '*.gz' -delete
