using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gerador
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGerar_Click(object sender, EventArgs e)
        {
            dynamic myObject = new ExpandoObject();
            myObject.numero_ordem_servico = txtOrdem.Text;
            myObject.data_de_emissao = txtDataEmissao.Text;
            myObject.nome = txtNome.Text;
            myObject.endereco = txtEndereco.Text;
            myObject.complemento = txtComplemento.Text;
            myObject.cidade = txtCidade.Text;
            myObject.cpf_cnpj = txtCPF.Text;
            myObject.bairro = txtBairro.Text;
            myObject.telefone = txtTelefone.Text;
            myObject.placa = txtPlaca.Text;
            myObject.renavam = txtRenavam.Text;
            myObject.chassi = txtChasi.Text;
            myObject.marca_modelo = txtMarcaModelo.Text;
            myObject.ano_fab_mob = txtAnoFab.Text;
            myObject.combustivel = txtComb.Text;
            myObject.cor = txtCor.Text;
            myObject.categoria = txtCat.Text;
            myObject.data_aquisicao = txtDataAq.Text;
            myObject.restricao = txtRest.Text;
            myObject.financeira = txtFinan.Text;
            myObject.valor_total = txtTotal.Text;
            myObject.adiantado = txtAdiantado.Text;
            myObject.devedor = txtDevedor.Text;

            List<dynamic> taxas = new List<dynamic>();
            var AllTaxas = txtTaxas.Text.Split('\n');
            foreach(var t in AllTaxas)
            {
                dynamic obj = new ExpandoObject();
                obj.taxa = t;
                taxas.Add(obj);
            }

            myObject.taxas = taxas;

            var dir = Directory.GetCurrentDirectory();
            var FilePath = dir + "/output.json";

            if (File.Exists(FilePath))
                File.Delete(FilePath);

            var jsonObj = JsonConvert.SerializeObject(myObject);
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter(FilePath, true))
            {
                sw.WriteLine(jsonObj);
                sw.Close();
            }

            System.Diagnostics.Process.Start(string.Format("{0}/generador.bat", dir));

            System.Threading.Thread.Sleep(4000);
            var FileDocxName = dir + "/filename.txt";
            var ContentDocxFile = File.ReadAllText(FileDocxName);
            File.Delete(FileDocxName);            
            System.Diagnostics.Process.Start(string.Format("{0}/{1}", dir, ContentDocxFile));
        }
    }
}
