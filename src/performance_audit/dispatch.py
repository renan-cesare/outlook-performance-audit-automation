import time
from datetime import datetime

import pandas as pd

from .excel_utils import read_excel_first_sheet, excel_is_locked
from .columns import find_column
from .token_utils import make_token
from .history_store import append_row
from .outlook_client import OutlookClient


def _clean_email(e: str) -> str:
    if e is None:
        return ""
    e = str(e).strip().replace("\u00a0", "").replace(" ", "")
    for ch in ["\u200b", "\u200c", "\u200d", "\ufeff"]:
        e = e.replace(ch, "")
    return e


def _email_ok(e: str) -> bool:
    import re
    return bool(re.match(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", _clean_email(e)))


def _norm_cod_assessor(raw) -> str:
    if pd.isna(raw):
        return ""
    s = str(raw).strip()
    if not s:
        return ""

    # trata casos de número no Excel (ex: 123.0)
    try:
        f = float(s.replace(",", "."))
        if f.is_integer():
            s = str(int(f))
    except Exception:
        pass

    s = s.upper()
    if not s.startswith("A"):
        s = "A" + s
    return s


def _load_html(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""


def _load_outlook_signature(signature_name: str) -> str:
    import os
    appdata = os.environ.get("APPDATA", "")
    sig_dir = os.path.join(appdata, "Microsoft", "Signatures")
    sig_path = os.path.join(sig_dir, f"{signature_name}.htm")
    return _load_html(sig_path) if os.path.exists(sig_path) else ""


def dispatch(cfg, logger):
    aud_path = cfg.get("paths.auditoria_xlsx")
    prof_path = cfg.get("paths.profissionais_xlsx")
    history_xlsx = cfg.get("paths.history_xlsx")
    sheet_name = cfg.get("history.sheet_name", "Performance_Audit_History")

    # bloqueio Excel aberto
    if excel_is_locked(aud_path):
        raise PermissionError(f"Feche a planilha de auditoria: {aud_path}")
    if excel_is_locked(prof_path):
        raise PermissionError(f"Feche a base de profissionais: {prof_path}")

    df_aud, sh_aud = read_excel_first_sheet(aud_path, None)
    df_prof, sh_prof = read_excel_first_sheet(prof_path, None)

    logger.info(f"Auditoria carregada | aba: {sh_aud} | linhas: {len(df_aud)}")
    logger.info(f"Profissionais carregada | aba: {sh_prof} | linhas: {len(df_prof)}")

    # ===== Colunas AUDITORIA (tolerante) =====
    c_cod_cli = find_column(df_aud, ["Cod Cliente", "Código Cliente", "Conta", "Codigo Cliente"])
    c_nome_cli = find_column(df_aud, ["Nome Cliente", "Cliente", "Nome do Cliente"])
    c_cod_ass = find_column(df_aud, ["Cod Assessor", "Código Assessor", "Codigo Assessor", "Assessor"])

    if not all([c_cod_cli, c_nome_cli, c_cod_ass]):
        raise RuntimeError(
            "Colunas obrigatórias não encontradas na auditoria. "
            f"cod_cliente={c_cod_cli}, nome_cliente={c_nome_cli}, cod_assessor={c_cod_ass}"
        )

    # ===== Colunas PROFISSIONAIS (tolerante) =====
    p_cod = find_column(df_prof, ["Cod Assessor", "Código", "Codigo", "Cod Profissional"])
    p_nome = find_column(df_prof, ["Nome Assessor", "Nome", "Assessor"])
    p_email = find_column(df_prof, ["E-mail", "Email", "Email Profissional", "E-mail Profissional"])

    p_cod_lider = find_column(df_prof, ["Cod Lider", "Código Líder", "Codigo Lider", "Cod Supervisor"])
    p_email_lider = find_column(df_prof, ["E-mail Líder", "Email Lider", "E-mail Supervisor", "Email Supervisor"])

    if not all([p_cod, p_nome, p_email]):
        raise RuntimeError(
            "Colunas obrigatórias não encontradas na base de profissionais. "
            f"cod={p_cod}, nome={p_nome}, email={p_email}"
        )

    # ===== Normaliza base de profissionais =====
    df_prof["_COD_NORM"] = df_prof[p_cod].apply(_norm_cod_assessor)
    df_prof = df_prof[df_prof["_COD_NORM"] != ""].copy()
    df_prof = df_prof.drop_duplicates(subset="_COD_NORM", keep="last")
    prof_map = df_prof.set_index("_COD_NORM").to_dict(orient="index")

    # ===== Limpa base de auditoria =====
    df_aud["_COD_ASS_NORM"] = df_aud[c_cod_ass].apply(_norm_cod_assessor)
    df_aud["_COD_CLI_STR"] = df_aud[c_cod_cli].astype(str).str.strip()
    df_aud = df_aud[(df_aud["_COD_ASS_NORM"] != "") & (df_aud["_COD_CLI_STR"].str.lower() != "nan")].copy()

    if df_aud.empty:
        logger.warn("Nenhuma linha válida na auditoria após limpeza.")
        return

    # ===== Outlook =====
    from_smtp = cfg.get("outlook.from_smtp")
    store_hint = cfg.get("outlook.store_hint", "riscos")
    oc = OutlookClient(from_smtp=from_smtp, store_hint=store_hint)

    sent_folder = oc.get_sent_folder() if cfg.get("outlook.use_sent_folder_override", True) else None

    # ===== Assinatura (opcional) =====
    signature_html = ""
    if cfg.get("signature.use_local_outlook_signature", True):
        signature_html = _load_outlook_signature(cfg.get("signature.signature_windows_name", ""))

    # ===== Template corpo =====
    body_template_path = cfg.get("paths.email_body_html")
    body_template = _load_html(body_template_path)

    # ===== Comportamento =====
    send_mode = str(cfg.get("behavior.send_mode", "display")).lower()   # display | send
    force_send = bool(cfg.get("behavior.force_send", True))
    delay = float(cfg.get("behavior.delay_between_emails_sec", 2.5))
    retry_send = int(cfg.get("behavior.retry_send", 2))
    max_emails = cfg.get("behavior.max_emails", None)

    # ===== Regras de e-mail =====
    month_ref = cfg.get("project.month_ref", "")
    subject_tpl = cfg.get("email.subject_template", "Auditoria – Cliente {nome_cliente} – {cod_cliente}")
    sla_days = cfg.get("email.sla_business_days", 3)
    skip_cc_codes = set(cfg.get("cc_rules.skip_cc_if_leader_in_codes", []))

    total = 0
    sent = 0
    skipped_no_prof = 0
    skipped_bad_email = 0

    for _, row in df_aud.iterrows():
        total += 1
        if max_emails and sent >= int(max_emails):
            logger.warn("Limite max_emails atingido. Encerrando.")
            break

        cod_cliente = str(row.get(c_cod_cli, "")).strip()
        nome_cliente = str(row.get(c_nome_cli, "")).strip()
        cod_ass = _norm_cod_assessor(row.get(c_cod_ass, ""))

        if cod_ass not in prof_map:
            skipped_no_prof += 1
            logger.warn(f"[PULADO] Cliente {cod_cliente}: assessor {cod_ass} não encontrado na base.")
            continue

        prof = prof_map[cod_ass]
        nome_assessor = str(prof.get(p_nome, "")).strip()
        to_email = _clean_email(prof.get(p_email, ""))

        if not _email_ok(to_email):
            skipped_bad_email += 1
            logger.warn(f"[PULADO] Cliente {cod_cliente}: assessor {cod_ass} sem e-mail válido.")
            continue

        # ===== CC do líder (mesma proposta do seu script) =====
        cc_email = ""
        cod_lider = _norm_cod_assessor(prof.get(p_cod_lider, "")) if p_cod_lider else ""

        if cod_lider and cod_lider in skip_cc_codes:
            cc_email = ""
        elif cod_lider and cod_lider in prof_map:
            cc_candidate = _clean_email(prof_map[cod_lider].get(p_email, ""))
            if _email_ok(cc_candidate):
                cc_email = cc_candidate
        elif p_email_lider:
            cc_candidate = _clean_email(prof.get(p_email_lider, ""))
            if _email_ok(cc_candidate):
                cc_email = cc_candidate

        token = make_token("PERF")
        subject = subject_tpl.format(nome_cliente=nome_cliente, cod_cliente=cod_cliente)

        # corpo com placeholders
        body = body_template.format(
            nome_assessor=nome_assessor,
            nome_cliente=nome_cliente,
            cod_cliente=cod_cliente,
            sla_business_days=sla_days
        )

        # token no corpo (fundamental pro follow-up)
        body += f"<br><br><p><small><b>Token:</b> {token}</small></p>"

        # assinatura (se existir)
        if signature_html:
            body += "<br><br>" + signature_html

        # ===== montar e-mail =====
        mail = oc.create_mail()
        mail.To = to_email
        if cc_email:
            mail.CC = cc_email
        mail.Subject = subject
        mail.HTMLBody = body

        if not mail.Recipients.ResolveAll():
            logger.warn(f"[PULADO] Cliente {cod_cliente}: destinatários não resolvidos.")
            continue

        # ===== display / send =====
        if send_mode == "display" or not force_send:
            try:
                mail.Save()
            except Exception:
                pass
            mail.Display()
            status = "PREPARADO"
        else:
            try:
                mail.Save()
            except Exception:
                pass

            last_ex = None
            ok = False
            for _ in range(retry_send):
                try:
                    mail.Send()
                    ok = True
                    break
                except Exception as ex:
                    last_ex = ex
                    time.sleep(1.2)

            if not ok:
                raise RuntimeError(f"Falha ao enviar (após {retry_send} tentativas): {last_ex}")

            status = "ENVIADO"
            sent += 1

            # move para "Enviados" da store alvo (se configurado)
            if sent_folder is not None:
                try:
                    mail.Move(sent_folder)
                except Exception:
                    pass

        ids = oc.extract_ids(mail)

        # ===== registrar histórico =====
        append_row(history_xlsx, sheet_name, {
            "datetime_sent": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "month_ref": month_ref,
            "cod_cliente": cod_cliente,
            "nome_cliente": nome_cliente,
            "cod_assessor": cod_ass,
            "nome_assessor": nome_assessor,
            "to_email": to_email,
            "cc_email": cc_email,
            "token": token,
            "subject": subject,
            "entry_id": ids.get("entry_id", ""),
            "conversation_id": ids.get("conversation_id", ""),
            "internet_message_id": ids.get("internet_message_id", ""),
            "status": status,
            "last_update_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "notes": ""
        })

        logger.info(f"[{status}] Cliente {cod_cliente} | {cod_ass} -> {to_email} (CC: {cc_email or '-'}) | token={token}")

        time.sleep(delay)

    logger.info("==== RESUMO DISPATCH ====")
    logger.info(f"Total processado: {total}")
    logger.info(f"Enviados: {sent}")
    logger.info(f"Pulado (sem assessor na base): {skipped_no_prof}")
    logger.info(f"Pulado (e-mail inválido): {skipped_bad_email}")

