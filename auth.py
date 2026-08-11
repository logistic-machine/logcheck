from __future__ import annotations

import hashlib
import hmac
import json
import os
from dataclasses import asdict, dataclass

from . import config


@dataclass
class Usuario:
    email: str
    nome: str
    senha_hash: str
    salt: str
    master: bool = False
    ativo: bool = False


# ---------------------------------------------------------------------------
# Senha (PBKDF2)
# ---------------------------------------------------------------------------

_ITERACOES = 200_000


def _hash(senha: str, salt: bytes) -> str:
    dk = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), salt, _ITERACOES)
    return dk.hex()


def gerar_credenciais(senha: str) -> tuple[str, str]:
    """Retorna (senha_hash, salt_hex) para uma senha em texto puro."""
    salt = os.urandom(16)
    return _hash(senha, salt), salt.hex()


def conferir_senha(senha: str, senha_hash: str, salt_hex: str) -> bool:
    try:
        salt = bytes.fromhex(salt_hex)
    except ValueError:
        return False
    return hmac.compare_digest(_hash(senha, salt), senha_hash)


# ---------------------------------------------------------------------------
# Persistência
# ---------------------------------------------------------------------------


def _normalizar_email(email: str) -> str:
    return email.strip().lower()


def carregar_usuarios() -> dict[str, Usuario]:
    """Carrega os usuários do JSON, criando os Masters iniciais se necessário."""
    usuarios: dict[str, Usuario] = {}
    if config.USUARIOS_JSON.exists():
        with open(config.USUARIOS_JSON, "r", encoding="utf-8") as fh:
            dados = json.load(fh)
        for reg in dados:
            u = Usuario(**reg)
            usuarios[u.email] = u

    # Garante a existência dos Masters iniciais (bootstrap).
    mudou = False
    for email in config.MASTER_EMAILS_INICIAIS:
        email = _normalizar_email(email)
        if email and email not in usuarios:
            senha_hash, salt = gerar_credenciais(config.SENHA_PADRAO)
            usuarios[email] = Usuario(
                email=email,
                nome=email.split("@")[0],
                senha_hash=senha_hash,
                salt=salt,
                master=True,
                ativo=True,
            )
            mudou = True
    if mudou:
        salvar_usuarios(usuarios)
    return usuarios


def salvar_usuarios(usuarios: dict[str, Usuario]) -> None:
    config.ensure_output_dirs()
    dados = [asdict(u) for u in usuarios.values()]
    with open(config.USUARIOS_JSON, "w", encoding="utf-8") as fh:
        json.dump(dados, fh, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Operações de alto nível
# ---------------------------------------------------------------------------


def autenticar(email: str, senha: str) -> Usuario | None:
    """Retorna o usuário se e-mail/senha conferirem; caso contrário, None."""
    usuarios = carregar_usuarios()
    u = usuarios.get(_normalizar_email(email))
    if u and conferir_senha(senha, u.senha_hash, u.salt):
        return u
    return None


def cadastrar_usuario(email: str, nome: str, senha: str | None = None,
                      master: bool = False, ativo: bool = True) -> Usuario:
    """Cria um novo usuário. Levanta ValueError se o e-mail já existir."""
    usuarios = carregar_usuarios()
    email = _normalizar_email(email)
    if not email or "@" not in email:
        raise ValueError("E-mail inválido.")
    if email in usuarios:
        raise ValueError("Já existe um usuário com esse e-mail.")
    senha_hash, salt = gerar_credenciais(senha or config.SENHA_PADRAO)
    u = Usuario(email=email, nome=nome.strip() or email.split("@")[0],
                senha_hash=senha_hash, salt=salt, master=master, ativo=ativo)
    usuarios[email] = u
    salvar_usuarios(usuarios)
    return u


def definir_ativo(email: str, ativo: bool) -> None:
    """Libera (SIM) ou bloqueia (NÃO) o acesso ao questionário."""
    usuarios = carregar_usuarios()
    email = _normalizar_email(email)
    if email in usuarios:
        usuarios[email].ativo = ativo
        salvar_usuarios(usuarios)


def definir_master(email: str, master: bool) -> None:
    """Concede ou retira o poder de Master de um e-mail."""
    usuarios = carregar_usuarios()
    email = _normalizar_email(email)
    if email in usuarios:
        usuarios[email].master = master
        if master:  # Master é sempre liberado.
            usuarios[email].ativo = True
        salvar_usuarios(usuarios)


def alterar_senha(email: str, nova_senha: str) -> None:
    usuarios = carregar_usuarios()
    email = _normalizar_email(email)
    if email in usuarios:
        senha_hash, salt = gerar_credenciais(nova_senha)
        usuarios[email].senha_hash = senha_hash
        usuarios[email].salt = salt
        salvar_usuarios(usuarios)
