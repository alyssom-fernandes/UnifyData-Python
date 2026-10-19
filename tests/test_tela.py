"""Teste da tela (unifydata.py) em modo exemplo, com o AppTest do Streamlit: sem navegador e sem enviar arquivo."""
import re
import unittest
from pathlib import Path

from streamlit.testing.v1 import AppTest

TELA = Path(__file__).resolve().parents[1] / 'unifydata.py'


def abrir_demo():
    at = AppTest.from_file(str(TELA), default_timeout=60)
    at.query_params['demo'] = '1'
    return at.run()


class Demonstracao(unittest.TestCase):
    def test_fluxo_completo_fecha_com_o_total_do_linx_menos_o_ignorado(self):
        at = abrir_demo()
        self.assertFalse(at.exception)
        rotulos = [b.label for b in at.button]
        self.assertEqual(rotulos.count('Aprovar vínculo'), 1, 'uma pessoa com nome parecido')
        self.assertEqual(rotulos.count('Ignorar'), 5, 'o nome parecido e as quatro fora do RH, cada um com o seu Ignorar')

        # Aprova o nome parecido e ignora as quatro pessoas que não estão no RH, como no roteiro da demonstração
        at.button[[b.label for b in at.button].index('Aprovar vínculo')].click().run()
        while any(b.label == 'Ignorar' for b in at.button):
            next(b for b in at.button if b.label == 'Ignorar').click().run()
        self.assertFalse(at.exception)

        texto = ' '.join(m.value for m in at.markdown)
        self.assertIn('Conferência com o Linx', texto)
        self.assertIn('Total geral no relatório', texto, 'a fita mostra o Total geral lido do relatório')
        self.assertEqual(re.findall(r'<dt>A descontar</dt><dd>([^<]+)</dd>', texto), ['R$ 2.159,00'])
        self.assertEqual(len(at.warning), 0, 'o exemplo fecha com o Total geral: nenhum aviso de divergência')

    def test_total_geral_diferente_mostra_aviso_e_marca_a_fita(self):
        at = abrir_demo()
        at.button[[b.label for b in at.button].index('Aprovar vínculo')].click().run()
        while any(b.label == 'Ignorar' for b in at.button):
            next(b for b in at.button if b.label == 'Ignorar').click().run()
        # Como se o relatório do Linx trouxesse mais 100,00 do que a leitura somou
        at.session_state['totais_relatorio'] = {**at.session_state['totais_relatorio'], 'LOJA': 1485.60}
        at.run()
        self.assertEqual(len(at.warning), 1)
        self.assertIn('não fecha com o Total geral', at.warning[0].value)
        self.assertIn('ud-fita-diverge', ' '.join(m.value for m in at.markdown))


if __name__ == '__main__':
    unittest.main()
