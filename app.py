from __future__ import annotations
import os
import subprocess
import sys
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import auth, config
from .data_loader import (
    ReferenciaCombinacoes,
    Secao,
    carregar_combinacoes,
    carregar_questionario,
    todas_perguntas,
)
from .results import Auditoria, salvar_auditoria, _send_email_relatorio

COR_PRIMARIA = "#3D8601"
COR_FUNDO = "#F4F6F9"


class LinhaPergunta:


    def __init__(self, master, pergunta, on_change):
        self.pergunta = pergunta
        self.on_change = on_change
        self.var = tk.StringVar(value="")
        self.evidencia_path: str | None = None

        self.frame = ttk.Frame(master, padding=(8, 6))
        self.frame.columnconfigure(0, weight=1)

        texto = f"{pergunta.codigo}  (peso {pergunta.peso:g})\n{pergunta.texto}"
        ttk.Label(self.frame, text=texto, wraplength=760, justify="left").grid(
            row=0, column=0, columnspan=4, sticky="w"
        )

        opcoes = ttk.Frame(self.frame)
        opcoes.grid(row=1, column=0, sticky="w", pady=(4, 0))
        for i, op in enumerate(config.RESPOSTAS):
            ttk.Radiobutton(
                opcoes, text=op, value=op, variable=self.var,
                command=self._resposta_mudou,
            ).grid(row=0, column=i, padx=(0, 14))

        # Área de anexo (aparece só quando resposta = NÃO).
        self.anexo_frame = ttk.Frame(self.frame)
        self.anexo_frame.grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.btn_anexo = ttk.Button(
            self.anexo_frame, text="Anexar evidência (PNG)", command=self._anexar
        )
        self.lbl_anexo = ttk.Label(self.anexo_frame, text="Nenhum arquivo", foreground="#B00020")
        self.btn_anexo.grid(row=0, column=0, padx=(0, 10))
        self.lbl_anexo.grid(row=0, column=1)

        self.lbl_detalhes = ttk.Label(self.anexo_frame, text="Justificativa")
        self.ent_detalhes = ttk.Entry(self.anexo_frame, width=60)
        self.lbl_detalhes.grid(row=1, column=0, pady=(8, 0), sticky="w")
        self.ent_detalhes.grid(row=1, column=1, columnspan=3, pady=(8, 0), sticky="w")
        self.anexo_frame.grid_remove()

        ttk.Separator(self.frame, orient="horizontal").grid(
            row=3, column=0, columnspan=4, sticky="ew", pady=(8, 0)
        )

    def _resposta_mudou(self):
        if self.var.get() == config.RESP_NAO:
            self.anexo_frame.grid()
        else:
            self.anexo_frame.grid_remove()
            self.evidencia_path = None
            self.lbl_anexo.config(text="Nenhum arquivo", foreground="#B00020")
            self.ent_detalhes.delete(0, "end")
        self.on_change()

    def _anexar(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a imagem de evidência (PNG)",
            filetypes=[("Imagem PNG", "*.png")],
        )
        if not caminho:
            return
        if not caminho.lower().endswith(".png"):
            messagebox.showerror("Formato inválido", "A evidência deve ser um arquivo PNG.")
            return
        self.evidencia_path = caminho
        self.lbl_anexo.config(text=Path(caminho).name, foreground="#137333")

    # API usada pela janela principal # 
    def resposta(self) -> str:
        return self.var.get()

    def pendente_anexo(self) -> bool:
        return self.var.get() == config.RESP_NAO and not self.evidencia_path

    def evidencia_detalhe(self) -> str:
        return self.ent_detalhes.get().strip()

    def marcar_sim(self):
        self.var.set(config.RESP_SIM)
        self.evidencia_path = None
        self.ent_detalhes.delete(0, "end")
        self.anexo_frame.grid_remove()
        self.lbl_anexo.config(text="Nenhum arquivo", foreground="#B00020")

    def limpar(self):
        self.var.set("")
        self.evidencia_path = None
        self.ent_detalhes.delete(0, "end")
        self.anexo_frame.grid_remove()
        self.lbl_anexo.config(text="Nenhum arquivo", foreground="#B00020")


class InventarioApp(tk.Tk):
    def __init__(self, secoes: list[Secao], referencia: ReferenciaCombinacoes,
                 usuario: auth.Usuario):
        super().__init__()
        self.secoes = secoes
        self.referencia = referencia
        self.usuario = usuario
        self.perguntas_widgets: list[LinhaPergunta] = []
        self.combo_vars: dict[str, tk.StringVar] = {}
        self.combos: dict[str, ttk.Combobox] = {}

        self.title("Desenvolvido por Gabriel Passos - Versão 1.0")
        self.geometry("980x760")
        self.configure(bg=COR_FUNDO)
        self._montar_estilos()
        self._montar_layout()
        self._atualizar_placar()

    # ------------------------------------------------------------------
    def _montar_estilos(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=COR_FUNDO)
        style.configure("TLabel", background=COR_FUNDO)
        style.configure("Header.TLabel", background=COR_PRIMARIA, foreground="white",
                        font=("Segoe UI", 15, "bold"), padding=10)
        style.configure("Sub.TLabel", font=("Segoe UI", 11, "bold"), foreground=COR_PRIMARIA)
        style.configure("Placar.TLabel", font=("Segoe UI", 11, "bold"), foreground=COR_PRIMARIA)
        style.configure("Salvar.TButton", font=("Segoe UI", 11, "bold"))

    def _montar_layout(self):
        ttk.Label(self, text="Checklist de Avaliação - Eneo Rio de Janeiro",
                  style="Header.TLabel", anchor="center").pack(fill="x")

        # Barra do usuário conectado #
        barra = ttk.Frame(self, padding=(12, 6))
        barra.pack(fill="x")
        marca = " · Master" if self.usuario.master else ""
        ttk.Label(barra,
                  text=f"Conectado: {self.usuario.nome} ({self.usuario.email}){marca}" #Tirar duvida se o responsavél é o amoxarife que está realizando a avaliação?#
                  ).pack(side="left")
        ttk.Button(barra, text="Sair", command=self._sair).pack(side="right")
        if self.usuario.master:
            ttk.Button(barra, text="Painel Master",
                       command=self._abrir_admin).pack(side="right", padx=(0, 8))

        # Área rolável com o formulário.
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=0, pady=0)
        canvas = tk.Canvas(container, bg=COR_FUNDO, highlightthickness=0)
        scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.form = ttk.Frame(canvas, padding=16)
        self.form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.form, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        self._montar_cabecalho(self.form)
        self._montar_parametros(self.form)
        self._montar_checklist(self.form)

        # Rodapé fixo com placar e ações.
        self._montar_rodape()

    # ------------------------------------------------------------------ 2.1
    def _montar_cabecalho(self, parent):
        ttk.Label(parent, text="1) Identificação", style="Sub.TLabel").pack(anchor="w")
        f = ttk.Frame(parent, padding=(0, 6, 0, 12))
        f.pack(fill="x")
        self.ent_responsavel = self._campo(f, "Responsável *", 0)
        self.email_responsavel = self._campo(f, "E-mail do Responsável *", 1)
        self.ent_data = self._campo(f, "Data *", 2)
        self.ent_data.insert(0, date.today().strftime("%d/%m/%Y"))
        self.ent_inspetor = self._campo(f, "Inspetor *", 3)
        self.ent_inspetor.insert(0, self.usuario.nome) 
    

    def _campo(self, parent, rotulo, col):
        wrap = ttk.Frame(parent)
        wrap.grid(row=0, column=col, padx=(0, 16), sticky="w")
        ttk.Label(wrap, text=rotulo).pack(anchor="w")
        ent = ttk.Entry(wrap, width=28)
        ent.pack()
        return ent

    # ------------------------------------------------------------------ 2.2
    def _montar_parametros(self, parent):
        ttk.Label(parent, text="2) Parâmetros do Centro",
                  style="Sub.TLabel").pack(anchor="w")
        f = ttk.Frame(parent, padding=(0, 6, 0, 12))
        f.pack(fill="x")
        for i, (campo, rotulo) in enumerate(config.COMBO_FIELDS):
            wrap = ttk.Frame(f)
            wrap.grid(row=i // 3, column=i % 3, padx=(0, 16), pady=4, sticky="w")
            ttk.Label(wrap, text=rotulo + " *").pack(anchor="w")
            var = tk.StringVar()
            cb = ttk.Combobox(wrap, textvariable=var, width=26, state="readonly")
            cb.pack()
            cb.bind("<<ComboboxSelected>>", lambda e, c=campo: self._combo_mudou(c))
            self.combo_vars[campo] = var
            self.combos[campo] = cb
        self._refrescar_combos()

    def _selecao(self) -> dict[str, str]:
        return {c: self.combo_vars[c].get() for c, _ in config.COMBO_FIELDS}

    def _combo_mudou(self, campo_alterado: str):
        # Limpa os campos posteriores ao alterado, evitando combinações inválidas.
        campos = [c for c, _ in config.COMBO_FIELDS]
        idx = campos.index(campo_alterado)
        for c in campos[idx + 1:]:
            self.combo_vars[c].set("")
        self._refrescar_combos()

    def _refrescar_combos(self):
        campos = [c for c, _ in config.COMBO_FIELDS]
        for i, campo in enumerate(campos):
            selecao_anterior = {c: self.combo_vars[c].get() for c in campos[:i]}
            opcoes = self.referencia.opcoes(campo, selecao_anterior)
            self.combos[campo]["values"] = opcoes
            # Autopreenche quando só há uma opção possível.
            atual = self.combo_vars[campo].get()
            if atual and atual not in opcoes:
                self.combo_vars[campo].set("")
            elif not atual and len(opcoes) == 1:
                self.combo_vars[campo].set(opcoes[0])

    # ------------------------------------------------------------------ 2.3
    def _montar_checklist(self, parent):
        topo = ttk.Frame(parent)
        topo.pack(fill="x")
        ttk.Label(topo, text="3) Checklist de Avaliação", style="Sub.TLabel").pack(side="left")
        for s in self.secoes:
            box = ttk.LabelFrame(parent, text=f"{s.numero}. {s.titulo}  (meta {s.meta:g})", padding=6)
            box.pack(fill="x", pady=(6, 4))
            for p in s.perguntas:
                lp = LinhaPergunta(box, p, self._atualizar_placar)
                lp.frame.pack(fill="x")
                self.perguntas_widgets.append(lp)

    # ------------------------------------------------------------------ 2.5
    def _montar_rodape(self):
        rod = ttk.Frame(self, padding=10)
        rod.pack(fill="x", side="bottom")
        self.lbl_placar = ttk.Label(rod, text="", style="Placar.TLabel")
        self.lbl_placar.pack(side="left")
        self.lbl_status = ttk.Label(rod, text="", wraplength=560)
        self.lbl_status.pack(side="left", padx=(16, 0))
        ttk.Button(rod, text="Salvar", style="Salvar.TButton",
                   command=self._salvar).pack(side="right", padx=(8, 0))
        ttk.Button(rod, text="Abrir resultados", command=self._abrir_resultados).pack(side="right", padx=(8, 0))
        ttk.Button(rod, text="Limpar", command=self._limpar).pack(side="right", padx=(8, 0))


    def _atualizar_placar(self):
        soma = 0.0
        qtd_sim = 0
        for lp in self.perguntas_widgets:
            if lp.resposta() == config.RESP_SIM:
                soma += lp.pergunta.peso
                qtd_sim += 1
        total = sum(p.peso for p in todas_perguntas(self.secoes))
        respondidas = sum(1 for lp in self.perguntas_widgets if lp.resposta())
        self.lbl_placar.config(
            text=(f"Respondidas: {respondidas}/{len(self.perguntas_widgets)}   |   "
                  f"Nota parcial: {soma:g} / {total:g}") 
        )

    # ------------------------------------------------------------------
    def _mostrar_status(self, mensagem: str, sucesso: bool = True):
        self.lbl_status.config(
            text=mensagem,
            foreground="#137333" if sucesso else "#B00020",
        )

    def _validar(self) -> str | None:
        if not self.ent_responsavel.get().strip():
            return "Informe o Responsável."
        if not self.ent_data.get().strip():
            return "Informe a Data."
        if not self.ent_inspetor.get().strip():
            return "Informe o Inspetor."
        email_responsavel = self.email_responsavel.get().strip()
        if not email_responsavel or "@" not in email_responsavel:
            return "Informe um e-mail de responsável válido."
        selecao = self._selecao()
        for campo, rotulo in config.COMBO_FIELDS:
            if not selecao[campo]:
                return f"Selecione o campo '{rotulo}'."
        if not self.referencia.combinacao_valida(selecao):
            return "A combinação de parâmetros selecionada não é válida."
        faltando = [lp.pergunta.codigo for lp in self.perguntas_widgets if not lp.resposta()]
        if faltando:
            return f"Responda todas as perguntas. Pendentes: {', '.join(faltando[:8])}..."
        sem_anexo = [lp.pergunta.codigo for lp in self.perguntas_widgets if lp.pendente_anexo()]
        if sem_anexo:
            return ("Anexe a evidência (PNG) para as respostas 'NÃO': "
                    + ", ".join(sem_anexo))
        return None

    def _salvar(self):
        erro = self._validar()
        if erro:
            self._mostrar_status(erro, False)
            messagebox.showwarning("Campos pendentes", erro)
            return
        selecao = self._selecao()
        respostas = {lp.pergunta.codigo: lp.resposta() for lp in self.perguntas_widgets}
        evidencias = {
            lp.pergunta.codigo: lp.evidencia_path
            for lp in self.perguntas_widgets
            if lp.evidencia_path
        }
        evidencias_detalhes = {
            lp.pergunta.codigo: lp.evidencia_detalhe()
            for lp in self.perguntas_widgets
            if lp.evidencia_detalhe()
        }
        aud = Auditoria(
            responsavel=self.ent_responsavel.get().strip(),
            data=self.ent_data.get().strip(),
            inspetor=self.ent_inspetor.get().strip(),
            area=selecao["area"], centro=selecao["centro"], empresa=selecao["empresa"],
            localidade=selecao["localidade"], tipo_centro=selecao["tipo_centro"],
            respostas=respostas, evidencias=evidencias,
            evidencia_detalhes=evidencias_detalhes,
        )
        try:
            res = salvar_auditoria(aud, self.secoes)
        except PermissionError:
            self._mostrar_status("Não foi possível salvar. Feche o arquivo Excel e tente novamente.", False)
            messagebox.showerror(
                "Arquivo aberto",
                "Não foi possível salvar. Feche o arquivo 'Inventario_Resultados.xlsx' "
                "no Excel e tente novamente.",
            )
            return
        except Exception as exc:  # noqa: BLE001 - feedback amigável ao usuário
            self._mostrar_status(f"Erro ao salvar: {exc}", False)
            messagebox.showerror("Erro ao salvar", str(exc))
            return

        try:
            destinatario = self.email_responsavel.get().strip()
            _send_email_relatorio(Path(res["arquivo"]), destinatario, aud, res["aba_detalhe"])
            self._mostrar_status(
                f"Auditoria {res['id']} salva com sucesso. Relatório enviado por e-mail.",
                True,
            )
        except Exception as exc:  # noqa: BLE001
            self._mostrar_status(
                f"Auditoria {res['id']} salva, mas não foi possível enviar o relatório por e-mail.",
                False,
            )
            messagebox.showwarning(
                "Relatório salvo",
                f"Auditoria salva, mas não foi possível enviar o relatório por e-mail: {exc}",
            )

        resumo = res["resumo"]
        messagebox.showinfo(
            "Auditoria salva",
            f"Auditoria {res['id']} registrada com sucesso!\n\n"
            f"Empresa: {aud.empresa}  |  Localidade: {aud.localidade}\n"
            f"Nota final: {resumo['nota_final']:g} de {resumo['peso_total']:g}\n"
            f"Aproveitamento: {resumo['aproveitamento']:.1f}%\n\n"
            f"Arquivo: {res['arquivo']}",
        )
        self._limpar()

    def _limpar(self):
        self.ent_responsavel.delete(0, "end")
        self.ent_inspetor.delete(0, "end")
        for var in self.combo_vars.values():
            var.set("")
        self._refrescar_combos()
        for lp in self.perguntas_widgets:
            lp.limpar()
        self._atualizar_placar()

    def _sair(self):
        self.destroy()

    def _abrir_admin(self):
        AdminUsuariosDialog(self, self.usuario)

    def _abrir_resultados(self):
        if not config.RESULTADOS_XLSX.exists():
            messagebox.showinfo("Sem resultados", "Nenhuma auditoria foi salva ainda.")
            return
        caminho = str(config.RESULTADOS_XLSX)
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", caminho])
            else:
                subprocess.Popen(["xdg-open", caminho])
        except Exception as exc:  # noqa: BLE001
            messagebox.showinfo("Resultados", f"Arquivo em:\n{caminho}\n\n({exc})")


class LoginDialog(tk.Tk):
    """Janela de acesso: e-mail corporativo + senha (mesma base da versão web)."""

    def __init__(self):
        super().__init__()
        self.usuario: auth.Usuario | None = None
        self.title("Desenvolvido por Gabriel Passos - Versão 1.0")
        self.geometry("440x300")
        self.configure(bg=COR_FUNDO)
        self._montar_estilos()

        ttk.Label(self, text="LogCheck",
                  style="Header.TLabel", anchor="center").pack(fill="x")
        corpo = ttk.Frame(self, padding=20)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text="Acesso", style="Sub.TLabel").pack(anchor="w", pady=(0, 8))

        ttk.Label(corpo, text="E-mail").pack(anchor="w")
        self.ent_email = ttk.Entry(corpo, width=40)
        self.ent_email.pack(fill="x")
        ttk.Label(corpo, text="Senha").pack(anchor="w", pady=(8, 0))
        self.ent_senha = ttk.Entry(corpo, width=40, show="*")
        self.ent_senha.pack(fill="x")
        self.ent_senha.bind("<Return>", lambda e: self._entrar())

        self.lbl_erro = ttk.Label(corpo, text="", foreground="#B00020")
        self.lbl_erro.pack(anchor="w", pady=(8, 0))
        ttk.Button(corpo, text="Entrar", style="Salvar.TButton",
                   command=self._entrar).pack(fill="x", pady=(10, 0))
        self.ent_email.focus_set()

    def _montar_estilos(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=COR_FUNDO)
        style.configure("TLabel", background=COR_FUNDO)
        style.configure("Header.TLabel", background=COR_PRIMARIA, foreground="white",
                        font=("Segoe UI", 14, "bold"), padding=10)
        style.configure("Sub.TLabel", font=("Segoe UI", 11, "bold"), foreground=COR_PRIMARIA)
        style.configure("Salvar.TButton", font=("Segoe UI", 11, "bold"))

    def _entrar(self):
        email = self.ent_email.get().strip()
        senha = self.ent_senha.get()
        u = auth.autenticar(email, senha)
        if u is None:
            self.lbl_erro.config(text="E-mail ou senha inválidos.")
            return
        if not u.ativo and not u.master:
            self.lbl_erro.config(
                text="Acesso ainda não liberado")
            return
        self.usuario = u
        self.destroy()


class AdminUsuariosDialog(tk.Toplevel):
    """Painel do Master: cadastrar, liberar/bloquear (SIM/NÃO) e gerir Masters."""

    def __init__(self, parent, usuario: auth.Usuario):
        super().__init__(parent)
        self.usuario = usuario
        self.title("Desenvolvido por Gabriel Passos - Versão 1.0")
        self.geometry("720x560")
        self.transient(parent)
        self.grab_set()

        cad = ttk.LabelFrame(self, text="Cadastrar novo usuário", padding=10)
        cad.pack(fill="x", padx=12, pady=(12, 6))
        linha = ttk.Frame(cad)
        linha.pack(fill="x")
        ttk.Label(linha, text="E-mail *").grid(row=0, column=0, sticky="w")
        self.ent_email = ttk.Entry(linha, width=30)
        self.ent_email.grid(row=1, column=0, padx=(0, 10))
        ttk.Label(linha, text="Nome").grid(row=0, column=1, sticky="w")
        self.ent_nome = ttk.Entry(linha, width=24)
        self.ent_nome.grid(row=1, column=1, padx=(0, 10))
        self.var_liberar = tk.BooleanVar(value=True)
        self.var_master = tk.BooleanVar(value=False)
        ttk.Checkbutton(linha, text="Liberar Acesso", variable=self.var_liberar).grid(
            row=1, column=2, padx=(0, 10))
        ttk.Checkbutton(linha, text="Tornar Master", variable=self.var_master).grid(
            row=1, column=3)
        ttk.Button(cad, text="Cadastrar", command=self._cadastrar).pack(
            anchor="e", pady=(8, 0))

        lista = ttk.LabelFrame(self, text="Usuários cadastrados", padding=10)
        lista.pack(fill="both", expand=True, padx=12, pady=(6, 6))
        cols = ("email", "nome", "acesso", "master")
        self.tree = ttk.Treeview(lista, columns=cols, show="headings", height=10)
        for c, t, w in (("email", "E-mail", 240), ("nome", "Nome", 160),
                        ("acesso", "Acesso", 90), ("master", "Master", 80)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(lista, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        acoes = ttk.Frame(self, padding=(12, 0, 12, 12))
        acoes.pack(fill="x")
        ttk.Button(acoes, text="Bloquear acesso",
                   command=self._alternar_ativo).pack(side="left")
        ttk.Button(acoes, text="Dar / Tirar Master",
                   command=self._alternar_master).pack(side="left", padx=8)
        ttk.Button(acoes, text="Resetar senha",
                   command=self._resetar_senha).pack(side="left")
        ttk.Button(acoes, text="Fechar", command=self.destroy).pack(side="right")

        self._recarregar()

    def _recarregar(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for email, u in sorted(auth.carregar_usuarios().items()):
            self.tree.insert("", "end", iid=email, values=(
                email, u.nome, "SIM" if u.ativo else "NÃO", "SIM" if u.master else "—"))

    def _selecionado(self) -> str | None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Selecione um almoxarife na lista.", parent=self)
            return None
        return sel[0]

    def _cadastrar(self):
        try:
            auth.cadastrar_usuario(
                self.ent_email.get(), self.ent_nome.get(),
                master=self.var_master.get(), ativo=self.var_liberar.get())
        except ValueError as e:
            messagebox.showerror("Erro", str(e), parent=self)
            return
        messagebox.showinfo(
            "Cadastrado",
            f"Almoxarife cadastrado (senha padrão: {config.SENHA_PADRAO}).", parent=self)
        self.ent_email.delete(0, "end")
        self.ent_nome.delete(0, "end")
        self._recarregar()

    def _alternar_ativo(self):
        email = self._selecionado()
        if not email:
            return
        if email == self.usuario.email:
            messagebox.showwarning("Ação inválida", "Você não pode bloquear a si mesmo.",
                                   parent=self)
            return
        u = auth.carregar_usuarios()[email]
        auth.definir_ativo(email, not u.ativo)
        self._recarregar()

    def _alternar_master(self):
        email = self._selecionado()
        if not email:
            return
        if email == self.usuario.email:
            messagebox.showwarning("Ação inválida", "Você não pode alterar seu próprio Master.",
                                   parent=self)
            return
        u = auth.carregar_usuarios()[email]
        auth.definir_master(email, not u.master)
        self._recarregar()

    def _resetar_senha(self):
        email = self._selecionado()
        if not email:
            return
        auth.alterar_senha(email, config.SENHA_PADRAO)
        messagebox.showinfo(
            "Senha redefinida",
            f"Senha de {email} redefinida para a padrão ({config.SENHA_PADRAO}).",
            parent=self)


def _erro_inicial(msg: str):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Erro ao iniciar", msg)
    root.destroy()


def main():
    try:
        combinacoes = carregar_combinacoes()
        secoes = carregar_questionario()
    except FileNotFoundError as exc:
        _erro_inicial(
            "Não foi possível encontrar as planilhas de dados.\n\n"
            f"{exc}\n\nVerifique a pasta 'data'."
        )
        return
    if not combinacoes or not secoes:
        _erro_inicial("As planilhas de dados estão vazias ou em formato inesperado.")
        return

    # Tela de acesso (mesma base de usuários da versão web).
    login = LoginDialog()
    login.mainloop()
    usuario = login.usuario
    if usuario is None:  # janela fechada sem login
        return

    app = InventarioApp(secoes, ReferenciaCombinacoes(combinacoes), usuario)
    app.mainloop()


if __name__ == "__main__":
    main()
