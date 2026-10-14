#!/bin/bash
# Instala o actualiza Liquidaciones y Fondos en el servidor (Ubuntu/Debian), detrás de nginx con HTTPS.
# Uso (como root, desde la carpeta del sistema):   bash servidor/instalar.sh [dominio] [puerto]
# Se puede volver a correr para actualizar: no toca datos.db ni la carpeta comprobantes.
set -euo pipefail

DOMINIO="${1:-liquidaciones.siderasrl.com.ar}"
PUERTO="${2:-5002}"
DIR=/opt/liquidaciones
ORIGEN="$(cd "$(dirname "$0")/.." && pwd)"
SERVICIO=liquidaciones

[ "$(id -u)" = 0 ] || { echo "Correlo como root (sudo bash servidor/instalar.sh)"; exit 1; }

echo "== Paquetes necesarios"
apt-get update -qq
apt-get install -y -qq python3 python3-venv nginx certbot python3-certbot-nginx rsync >/dev/null

if ss -ltn | grep -q ":$PUERTO " && ! systemctl is-active --quiet "$SERVICIO"; then
  echo "El puerto $PUERTO ya lo usa otro programa. Volvé a correr con otro puerto, ej: bash servidor/instalar.sh $DOMINIO 5003"
  exit 1
fi

echo "== Copiando el sistema a $DIR (sin tocar datos.db ni comprobantes)"
id -u liquidaciones >/dev/null 2>&1 || useradd --system --home "$DIR" --shell /usr/sbin/nologin liquidaciones
mkdir -p "$DIR/comprobantes"
if [ "$ORIGEN" != "$DIR" ]; then
  rsync -a --delete --exclude datos.db --exclude comprobantes/ --exclude .git/ --exclude venv/ \
        --exclude '*.xlsx' --exclude __pycache__/ "$ORIGEN"/ "$DIR"/
  # Primera instalación desde el zip con datos: se copia la base solo si el servidor todavía no tiene una
  if [ -f "$ORIGEN/datos.db" ] && [ ! -f "$DIR/datos.db" ]; then
    echo "== Copiando la base de datos incluida (primera instalación)"
    cp "$ORIGEN/datos.db" "$DIR/datos.db"
    [ -d "$ORIGEN/comprobantes" ] && cp -r "$ORIGEN/comprobantes/." "$DIR/comprobantes/"
  fi
fi
python3 -m venv "$DIR/venv"
"$DIR/venv/bin/pip" install -q --upgrade pip
"$DIR/venv/bin/pip" install -q -r "$DIR/requirements.txt"
chown -R liquidaciones:liquidaciones "$DIR"
chmod 750 "$DIR"

echo "== Servicio $SERVICIO (arranca solo con el servidor)"
cat > /etc/systemd/system/$SERVICIO.service <<UNIT
[Unit]
Description=Liquidaciones y Fondos (SIDERA)
After=network.target

[Service]
User=liquidaciones
WorkingDirectory=$DIR
Environment=LF_HOST=127.0.0.1
Environment=LF_PORT=$PUERTO
Environment=LF_HTTPS=1
Environment=LF_NO_BROWSER=1
ExecStart=$DIR/venv/bin/python app.py
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
UNIT
systemctl daemon-reload
systemctl enable --now $SERVICIO
systemctl restart $SERVICIO

echo "== nginx para $DOMINIO"
SITIO=/etc/nginx/sites-available/$SERVICIO
if [ ! -f "$SITIO" ]; then
cat > "$SITIO" <<NGINX
server {
    listen 80;
    server_name $DOMINIO;
    client_max_body_size 30m;
    location / {
        proxy_pass http://127.0.0.1:$PUERTO;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        proxy_set_header X-Forwarded-Host \$host;
    }
}
NGINX
ln -sf "$SITIO" /etc/nginx/sites-enabled/$SERVICIO
fi
nginx -t
systemctl reload nginx

echo "== Certificado HTTPS"
if ! certbot certificates 2>/dev/null | grep -q "Domains:.*$DOMINIO"; then
  certbot --nginx -d "$DOMINIO" --non-interactive --agree-tos --redirect --register-unsafely-without-email \
    || { echo "No se pudo sacar el certificado. Revisá que $DOMINIO ya apunte a este servidor (registro DNS tipo A) y volvé a correr el script."; exit 1; }
fi

sleep 2
if systemctl is-active --quiet $SERVICIO; then
  echo
  echo "Listo: https://$DOMINIO"
  echo "Si la base está vacía, la primera vez pide crear el usuario administrador."
else
  echo "El servicio no arrancó. Ver el error con: journalctl -u $SERVICIO -n 50"
  exit 1
fi
