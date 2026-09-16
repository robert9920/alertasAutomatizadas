"""Armado del mensaje MIME y envio por SMTP."""
from __future__ import annotations

import mimetypes
import smtplib
import ssl
from dataclasses import dataclass, field
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

from .email_builder import CID_GRAFICO


class ErrorEnvio(Exception):
    """Fallo controlado durante el envio del correo."""


@dataclass
class Envio:
    asunto: str = ""
    html: str = ""
    texto: str = ""
    imagen: bytes | None = None
    para: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    cco: list[str] = field(default_factory=list)
    adjuntos: list[Path] = field(default_factory=list)

    @property
    def destinos(self) -> list[str]:
        vistos: list[str] = []
        for correo in self.para + self.cc + self.cco:
            limpio = correo.strip()
            if limpio and limpio.lower() not in [v.lower() for v in vistos]:
                vistos.append(limpio)
        return vistos

    def problemas(self) -> list[str]:
        faltas = []
        if not self.para:
            faltas.append("No hay ningún destinatario en 'Para'.")
        if not self.asunto.strip():
            faltas.append("El asunto está vacío.")
        if not self.html.strip():
            faltas.append("El cuerpo del correo está vacío.")
        for adjunto in self.adjuntos:
            if not Path(adjunto).is_file():
                faltas.append(f"No se encuentra el adjunto: {adjunto}")
        return faltas


# --------------------------------------------------------------------------- #
def construir_mensaje(cfg, envio: Envio) -> EmailMessage:
    mensaje = EmailMessage()
    mensaje["Subject"] = envio.asunto
    mensaje["From"] = (
        formataddr((cfg.nombre_remitente, cfg.desde)) if cfg.nombre_remitente else cfg.desde
    )
    if envio.para:
        mensaje["To"] = ", ".join(envio.para)
    if envio.cc:
        mensaje["Cc"] = ", ".join(envio.cc)
    if cfg.responder_a:
        mensaje["Reply-To"] = cfg.responder_a
    mensaje["Date"] = formatdate(localtime=True)
    mensaje["Message-ID"] = make_msgid()

    mensaje.set_content(envio.texto or "Este correo requiere un lector con HTML.")
    mensaje.add_alternative(envio.html, subtype="html")

    if envio.imagen:
        parte_html = mensaje.get_payload()[-1]
        parte_html.add_related(
            envio.imagen, maintype="image", subtype="png",
            cid=f"<{CID_GRAFICO}>", filename="curva_s.png",
        )

    for adjunto in envio.adjuntos:
        ruta = Path(adjunto)
        tipo, _ = mimetypes.guess_type(ruta.name)
        principal, _, secundario = (tipo or "application/octet-stream").partition("/")
        mensaje.add_attachment(
            ruta.read_bytes(), maintype=principal, subtype=secundario or "octet-stream",
            filename=ruta.name,
        )
    return mensaje


def _conectar(cfg):
    contexto = ssl.create_default_context()
    if cfg.seguridad == "SSL":
        servidor = smtplib.SMTP_SSL(cfg.servidor, cfg.puerto, timeout=cfg.timeout,
                                    context=contexto)
    else:
        servidor = smtplib.SMTP(cfg.servidor, cfg.puerto, timeout=cfg.timeout)
        servidor.ehlo()
        if cfg.seguridad == "STARTTLS":
            servidor.starttls(context=contexto)
            servidor.ehlo()
    if cfg.usuario and cfg.contrasena:
        servidor.login(cfg.usuario, cfg.contrasena)
    return servidor


def _traducir(exc: Exception) -> str:
    if isinstance(exc, smtplib.SMTPAuthenticationError):
        return (
            "El servidor rechazó el usuario o la contraseña.\n\n"
            "En Microsoft 365 suele deberse a que la autenticación SMTP está "
            "deshabilitada o a que la cuenta tiene MFA: pide a TI que habilite "
            "SMTP AUTH o genera una contraseña de aplicación.\n\n"
            f"Detalle: {exc}"
        )
    if isinstance(exc, smtplib.SMTPRecipientsRefused):
        return f"El servidor rechazó estos destinatarios: {exc.recipients}"
    if isinstance(exc, smtplib.SMTPSenderRefused):
        return f"El servidor rechazó el remitente indicado.\n\nDetalle: {exc}"
    if isinstance(exc, (TimeoutError, OSError)):
        return (
            "No se pudo conectar con el servidor SMTP. Revisa el servidor, el "
            f"puerto y la conexión de red.\n\nDetalle: {exc}"
        )
    return f"No se pudo enviar el correo.\n\nDetalle: {exc}"


def probar_conexion(cfg) -> str:
    faltas = cfg.problemas()
    if faltas:
        raise ErrorEnvio("Faltan datos en la hoja SMTP: " + ", ".join(faltas))
    try:
        servidor = _conectar(cfg)
    except Exception as exc:                                    # noqa: BLE001
        raise ErrorEnvio(_traducir(exc)) from exc
    try:
        servidor.noop()
    finally:
        try:
            servidor.quit()
        except Exception:                                       # noqa: BLE001
            pass
    return f"Conexión correcta con {cfg.servidor}:{cfg.puerto} ({cfg.seguridad})."


def enviar(cfg, envio: Envio) -> None:
    faltas = cfg.problemas()
    if faltas:
        raise ErrorEnvio("Faltan datos en la hoja SMTP: " + ", ".join(faltas))
    faltas = envio.problemas()
    if faltas:
        raise ErrorEnvio("\n".join(faltas))

    mensaje = construir_mensaje(cfg, envio)
    try:
        servidor = _conectar(cfg)
    except Exception as exc:                                    # noqa: BLE001
        raise ErrorEnvio(_traducir(exc)) from exc
    try:
        servidor.send_message(mensaje, from_addr=cfg.desde, to_addrs=envio.destinos)
    except Exception as exc:                                    # noqa: BLE001
        raise ErrorEnvio(_traducir(exc)) from exc
    finally:
        try:
            servidor.quit()
        except Exception:                                       # noqa: BLE001
            pass


def guardar_eml(cfg, envio: Envio, destino: Path) -> Path:
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    mensaje = construir_mensaje(cfg, envio)
    destino.write_bytes(mensaje.as_bytes())
    return destino
