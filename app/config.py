from __future__ import annotations

import os

from dataclasses import dataclass


@dataclass(frozen=True)
class PortalConfig:
    login_url: str = os.environ.get("INVOICE_LOGIN_URL", "https://portal.example.invalid/")
    list_url: str = os.environ.get("INVOICE_LIST_URL", "https://portal.example.invalid/")


@dataclass(frozen=True)
class SheetExcelConfig:
    sheet_name: str
    col_invoice_aliases: tuple[str, ...]
    col_emissao_fatura_aliases: tuple[str, ...]
    col_status_aliases: tuple[str, ...]
    col_motivo_aliases: tuple[str, ...]


@dataclass(frozen=True)
class ExcelConfig:
    value_not_found: str = "Não encontrado"

    sheets: tuple[SheetExcelConfig, ...] = (
        SheetExcelConfig(
            sheet_name="Vencidos",
            col_invoice_aliases=(
                "Nº FATURA",
                "N° FATURA",
                "N FATURA",
                "NUM FATURA",
                "NUMERO FATURA",
                "NÚMERO FATURA",
            ),
            col_emissao_fatura_aliases=(
                "EMISSÃO FATURA",
                "EMISSAO FATURA",
                "DATA EMISSÃO FATURA",
                "DATA EMISSAO FATURA",
            ),
            col_status_aliases=(
                "Situação fatura no Transpofrete",
                "Situacao fatura no Transpofrete",
            ),
            col_motivo_aliases=(
                "Motivo rejeite no Transpofrete",
                "Motivo rejeição no Transpofrete",
                "Motivo rejeicao no Transpofrete",
            ),
        ),
        SheetExcelConfig(
            sheet_name="A Vencer",
            col_invoice_aliases=(
                "Nº FATURA",
                "N° FATURA",
                "N FATURA",
                "NUM FATURA",
                "NUMERO FATURA",
                "NÚMERO FATURA",
            ),
            col_emissao_fatura_aliases=(
                "EMISSÃO FATURA",
                "EMISSAO FATURA",
                "DATA EMISSÃO FATURA",
                "DATA EMISSAO FATURA",
            ),
            col_status_aliases=(
                "Status TSP",
                "STATUS TSP",
                "Status Transpofrete",
                "Status no Transpofrete",
            ),
            col_motivo_aliases=(
                "OBS.",
                "OBS",
                "Observação",
                "Observacao",
                "Observações",
                "Observacoes",
            ),
        ),
    )


@dataclass(frozen=True)
class Selectors:
    login_input: str = '[id="frmLogin:txtLogin"]'
    password_input: str = '[id="frmLogin:txtSenha"]'

    date_from: str = '[id="form:dateDedataEmissao_input"]'
    date_to: str = '[id="form:dateAtedataEmissao_input"]'

    table_body: str = '[id$=":listaTable_data"]'
    table_rows: str = '[id$=":listaTable_data"] tr'

    reason_textarea: str = (
        'textarea[id$="motivoRejeicao"], '
        'textarea[name$="motivoRejeicao"], '
        'textarea[readonly]'
    )


PORTAL = PortalConfig()
EXCEL = ExcelConfig()
SELECTORS = Selectors()