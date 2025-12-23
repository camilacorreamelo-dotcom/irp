import os
import tempfile

import pandas as pd
import streamlit as st

import rodar_aghu
import rodar_irp


def main():
    st.title("Automação AGHU + ComprasNet (IRP)")

    # ==========================================
    # 1. Upload da planilha de entrada
    # ==========================================
    st.markdown("### 1. Planilha de entrada (AGHU)")
    uploaded_file = st.file_uploader(
        "Envie a planilha de entrada (Excel) com os códigos AGHU",
        type=["xlsx", "xls"],
        help="É a planilha base com Código AGHU, Código CATMAT, descrição, quantidade etc.",
    )

    # ==========================================
    # 2. Login AGHU
    # ==========================================
    st.markdown("### 2. Login AGHU")
    aghu_usuario = st.text_input("Usuário AGHU")
    aghu_senha = st.text_input("Senha AGHU", type="password")

    # ==========================================
    # 3. Login ComprasNet
    # ==========================================
    st.markdown("### 3. Login ComprasNet")
    cpf_compras = st.text_input("CPF (ComprasNet)")
    senha_compras = st.text_input("Senha ComprasNet", type="password")

    # ==========================================
    # 4. Número da IRP
    # ==========================================
    st.markdown("### 4. IRP a ser atualizada")
    irp_numero = st.text_input("Número da IRP (ex: 155022 - 00080/2025)")

    # ==========================================
    # Botão principal
    # ==========================================
    if st.button("🚀 Executar automação completa"):
        # validações simples
        if uploaded_file is None:
            st.error("Por favor, envie a planilha de entrada (AGHU) antes de executar.")
            return

        if not aghu_usuario or not aghu_senha:
            st.error("Preencha usuário e senha do AGHU.")
            return

        if not cpf_compras or not senha_compras:
            st.error("Preencha CPF e senha do ComprasNet.")
            return

        if not irp_numero:
            st.error("Informe o número da IRP que deseja atualizar.")
            return

        # ======================================================
        # 1) Salva a planilha enviada num arquivo temporário
        # ======================================================
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
            tmp_in.write(uploaded_file.getbuffer())
            excel_entrada_path = tmp_in.name

        tmp_dir = os.path.dirname(excel_entrada_path)

        # Arquivos intermediários (uso interno)
        excel_aghu_saida_path = os.path.join(tmp_dir, "AGHU_CONSUMO_ATUALIZADO_tmp.xlsx")
        relatorio_irp_path = os.path.join(tmp_dir, "RELATORIO_IRP_RESULTADO_tmp.xlsx")

        # Arquivo final único que o usuário vai baixar
        relatorio_final_path = os.path.join(tmp_dir, "RESULTADO_FINAL_AGHU_IRP.xlsx")

        # Aponta caminhos nos módulos
        rodar_aghu.EXCEL_ENTRADA = excel_entrada_path
        rodar_aghu.EXCEL_SAIDA = excel_aghu_saida_path

        rodar_irp.EXCEL_PATH = excel_aghu_saida_path
        rodar_irp.RELATORIO_SAIDA = relatorio_irp_path

        st.info(
            "Os logins no AGHU e no ComprasNet serão feitos automaticamente com as credenciais informadas acima. "
            "Não feche a janela do navegador enquanto a automação estiver rodando."
        )

        try:
            # ===========================
            # 2) Etapa AGHU
            # ===========================
            st.markdown("### Etapa 1/2 — Consultando preços no AGHU...")
            rodar_aghu.rodar_aghu(aghu_usuario, aghu_senha)
            st.success("AGHU finalizado. Planilha com preços atualizada internamente.")

            # ===========================
            # 3) Etapa IRP
            # ===========================
            st.markdown("### Etapa 2/2 — Lançando itens na IRP (ComprasNet)...")
            # 👇 AQUI PASSAMOS O NÚMERO DA IRP
            rodar_irp.rodar_irp(
                cpf=cpf_compras,
                senha=senha_compras,
                irp_numero=irp_numero,
            )
            st.success("IRP finalizada. Relatório interno de status/motivo gerado.")

            # ===========================
            # 4) Consolidar tudo em UMA planilha final
            # ===========================
            st.markdown("### Consolidando resultados em uma única planilha...")

            # lê planilha do AGHU (preço + 'Encontrado no AGHU')
            df_aghu = pd.read_excel(excel_aghu_saida_path)

            # lê relatório da IRP (status/motivo por linha_excel)
            df_irp = pd.read_excel(relatorio_irp_path)

            # garante colunas que vamos usar
            for col in ["linha_excel", "status", "motivo"]:
                if col not in df_irp.columns:
                    df_irp[col] = ""

            # adiciona coluna linha_excel na planilha do AGHU (é o índice original)
            df_aghu_reset = df_aghu.reset_index().rename(columns={"index": "linha_excel"})

            # seleciona só o que interessa da IRP e renomeia colunas
            df_irp_status = df_irp[["linha_excel", "status", "motivo"]].rename(
                columns={
                    "status": "Status IRP",
                    "motivo": "Motivo IRP",
                }
            )

            # merge: AGHU + info da IRP
            df_final = df_aghu_reset.merge(df_irp_status, on="linha_excel", how="left")

            # se você não quiser ver a coluna de controle, pode remover:
            df_final = df_final.drop(columns=["linha_excel"])

            # salva apenas o arquivo final
            df_final.to_excel(relatorio_final_path, index=False)

            st.success("✅ Planilha final única gerada com sucesso!")

            # ===========================
            # 5) Download do arquivo único
            # ===========================
            try:
                with open(relatorio_final_path, "rb") as f:
                    st.download_button(
                        "⬇️ Baixar RESULTADO_FINAL_AGHU_IRP.xlsx",
                        data=f,
                        file_name="RESULTADO_FINAL_AGHU_IRP.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except FileNotFoundError:
                st.warning("Não consegui localizar o arquivo final para download.")

        except Exception as e:
            st.error(f"⚠️ Ocorreu um erro durante a automação: {e}")


if __name__ == "__main__":
    main()
