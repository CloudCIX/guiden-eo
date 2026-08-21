FROM ghcr.io/euro-office/documentserver@sha256:6068075cc247f9283d337b2cf9a105d9df49fcd8cd6b5e82fa90640df1552e80

COPY runtime/v1/ \
    /var/www/euro-office/documentserver/sdkjs-plugins/v1/

COPY plugin/cloudcix-guiden/ \
    /var/www/euro-office/documentserver/sdkjs-plugins/cloudcix-guiden/
