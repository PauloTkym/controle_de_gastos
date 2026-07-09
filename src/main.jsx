import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  ArrowDownCircle,
  ArrowUpCircle,
  BarChart3,
  CalendarDays,
  Download,
  Plus,
  Repeat2,
  RefreshCw,
  Trash2,
  UserPlus,
  Wallet,
} from 'lucide-react';
import './styles.css';

if ('serviceWorker' in navigator && import.meta.env.PROD) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register(`${import.meta.env.BASE_URL}service-worker.js`);
  });
}

const STORAGE_KEY = 'controle-gastos-react-lancamentos';
const PESSOAS_STORAGE_KEY = 'controle-gastos-react-pessoas';
const CATEGORIAS_STORAGE_KEY = 'controle-gastos-react-categorias';
const PAGAMENTOS_STORAGE_KEY = 'controle-gastos-react-pagamentos';
const RECORRENTES_STORAGE_KEY = 'controle-gastos-react-recorrentes';

const categoriasPadrao = {
  Saida: ['Alimentacao', 'Transporte', 'Moradia', 'Lazer', 'Saude', 'Educacao', 'Outros'],
  Entrada: ['Salario', 'Aluguel', 'Investimentos', 'Outros'],
};

const labelsCategorias = {
  Alimentacao: 'Alimentacao',
  Transporte: 'Transporte',
  Moradia: 'Moradia',
  Lazer: 'Lazer',
  Saude: 'Saude',
  Educacao: 'Educacao',
  Outros: 'Outros',
  Salario: 'Salario',
  Aluguel: 'Aluguel',
  Investimentos: 'Investimentos',
};

const pagamentosPadrao = {
  Saida: ['Dinheiro', 'Cartao de Credito', 'Cartao de Debito', 'Pix', 'Boleto', 'Verocard'],
  Entrada: ['Deposito Bancario', 'Pix', 'Dinheiro', 'Boleto'],
};

const labelsPagamentos = {
  Dinheiro: 'Dinheiro',
  'Cartao de Credito': 'Cartao de Credito',
  'Cartao de Debito': 'Cartao de Debito',
  Pix: 'Pix',
  Boleto: 'Boleto',
  Verocard: 'Verocard',
  'Deposito Bancario': 'Deposito Bancario',
};

function hojeISO() {
  const hoje = new Date();
  return `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}-${String(
    hoje.getDate(),
  ).padStart(2, '0')}`;
}

function fimDoMesAtualISO() {
  const hoje = new Date();
  const fimDoMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
  return dataParaISO(fimDoMes);
}

function mesAtual() {
  return hojeISO().slice(0, 7);
}

function carregarLancamentos() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
  } catch {
    return [];
  }
}

function carregarPessoas() {
  try {
    return JSON.parse(localStorage.getItem(PESSOAS_STORAGE_KEY)) || [];
  } catch {
    return [];
  }
}

function carregarRecorrentes() {
  try {
    return JSON.parse(localStorage.getItem(RECORRENTES_STORAGE_KEY)) || [];
  } catch {
    return [];
  }
}

function carregarOpcoes(chave, padrao) {
  try {
    const salvas = JSON.parse(localStorage.getItem(chave));
    return salvas || padrao;
  } catch {
    return padrao;
  }
}

function salvarLancamentos(lancamentos) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(lancamentos));
}

function salvarPessoas(pessoas) {
  localStorage.setItem(PESSOAS_STORAGE_KEY, JSON.stringify(pessoas));
}

function salvarOpcoes(chave, opcoes) {
  localStorage.setItem(chave, JSON.stringify(opcoes));
}

function salvarRecorrentes(recorrentes) {
  localStorage.setItem(RECORRENTES_STORAGE_KEY, JSON.stringify(recorrentes));
}

function formatarMoeda(valor) {
  return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function formatarData(dataISO) {
  if (!dataISO) return 'Sem data';
  const [ano, mes, dia] = dataISO.split('-');
  return `${dia}/${mes}/${ano}`;
}

function formatarMes(mesISO) {
  if (mesISO === 'todos') return 'Todos os meses';
  const [ano, mes] = mesISO.split('-');
  return new Date(Number(ano), Number(mes) - 1, 1).toLocaleDateString('pt-BR', {
    month: 'long',
    year: 'numeric',
  });
}

function tipoDaTransacao(item) {
  return item.tipo || (item.valor >= 0 ? 'Entrada' : 'Saida');
}

function opcoesFiltroPorTipo(tipoFiltro, opcoes, transacoes, campo) {
  const opcoesCadastradas =
    tipoFiltro === 'todos' ? Object.values(opcoes).flat() : opcoes[tipoFiltro] || [];
  const opcoesUsadas = transacoes
    .filter((item) => tipoFiltro === 'todos' || tipoDaTransacao(item) === tipoFiltro)
    .map((item) => item[campo]);

  return [...new Set([...opcoesCadastradas, ...opcoesUsadas].filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, 'pt-BR'),
  );
}

function normalizarValor(valorTexto) {
  const texto = String(valorTexto).trim().replace(/\s/g, '').replace(',', '.');
  const numero = Number.parseFloat(texto);
  return Number.isFinite(numero) ? numero : 0;
}

function dataLocal(dataISO) {
  const [ano, mes, dia] = dataISO.split('-').map(Number);
  return new Date(ano, mes - 1, dia);
}

function dataParaISO(data) {
  return `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, '0')}-${String(data.getDate()).padStart(2, '0')}`;
}

function diasNoMes(ano, mesZeroBased) {
  return new Date(ano, mesZeroBased + 1, 0).getDate();
}

function somarMeses(data, quantidade, diaReferencia = data.getDate()) {
  const destino = new Date(data.getFullYear(), data.getMonth() + quantidade, 1);
  const diaFinal = Math.min(diaReferencia, diasNoMes(destino.getFullYear(), destino.getMonth()));
  destino.setDate(diaFinal);
  return destino;
}

function proximaDataRecorrente(data, intervalo, unidade, diaReferencia) {
  const proxima = new Date(data);
  if (unidade === 'dias') {
    proxima.setDate(proxima.getDate() + intervalo);
    return proxima;
  }
  if (unidade === 'semanas') {
    proxima.setDate(proxima.getDate() + intervalo * 7);
    return proxima;
  }
  if (unidade === 'meses') {
    return somarMeses(data, intervalo, diaReferencia);
  }
  return somarMeses(data, intervalo * 12, diaReferencia);
}

function App() {
  const [abaAtiva, setAbaAtiva] = useState('form');
  const [tipo, setTipo] = useState('Saida');
  const [data, setData] = useState(hojeISO());
  const [pessoa, setPessoa] = useState('');
  const [novoNome, setNovoNome] = useState('');
  const [categoria, setCategoria] = useState('');
  const [novaCategoria, setNovaCategoria] = useState('');
  const [descricao, setDescricao] = useState('');
  const [formaPagamento, setFormaPagamento] = useState('');
  const [novaFormaPagamento, setNovaFormaPagamento] = useState('');
  const [valor, setValor] = useState('');
  const [tipoRecorrente, setTipoRecorrente] = useState('Saida');
  const [dataInicioRecorrente, setDataInicioRecorrente] = useState(hojeISO());
  const [dataFimRecorrente, setDataFimRecorrente] = useState('');
  const [pessoaRecorrente, setPessoaRecorrente] = useState('');
  const [categoriaRecorrente, setCategoriaRecorrente] = useState('');
  const [descricaoRecorrente, setDescricaoRecorrente] = useState('');
  const [formaPagamentoRecorrente, setFormaPagamentoRecorrente] = useState('');
  const [valorRecorrente, setValorRecorrente] = useState('');
  const [intervaloRecorrente, setIntervaloRecorrente] = useState('1');
  const [unidadeRecorrente, setUnidadeRecorrente] = useState('meses');
  const [vezesRecorrente, setVezesRecorrente] = useState('');
  const [alerta, setAlerta] = useState(null);
  const [lancamentos, setLancamentos] = useState(carregarLancamentos);
  const [recorrentes, setRecorrentes] = useState(carregarRecorrentes);
  const [pessoas, setPessoas] = useState(carregarPessoas);
  const [categorias, setCategorias] = useState(() => carregarOpcoes(CATEGORIAS_STORAGE_KEY, categoriasPadrao));
  const [pagamentos, setPagamentos] = useState(() => carregarOpcoes(PAGAMENTOS_STORAGE_KEY, pagamentosPadrao));
  const [mesFiltro, setMesFiltro] = useState(mesAtual());
  const [pessoaFiltro, setPessoaFiltro] = useState('todos');
  const [tipoFiltro, setTipoFiltro] = useState('todos');
  const [categoriaFiltro, setCategoriaFiltro] = useState('todos');
  const [pagamentoFiltro, setPagamentoFiltro] = useState('todos');

  const ocorrenciasRecorrentes = useMemo(() => gerarOcorrenciasRecorrentes(recorrentes), [recorrentes]);

  const transacoes = useMemo(() => {
    return [...lancamentos, ...ocorrenciasRecorrentes];
  }, [lancamentos, ocorrenciasRecorrentes]);

  const resumoGeral = useMemo(() => calcularResumo(transacoes), [transacoes]);

  const mesesDisponiveis = useMemo(() => {
    const meses = new Set(transacoes.map((item) => item.data?.slice(0, 7)).filter(Boolean));
    meses.add(mesAtual());
    return [...meses].sort((a, b) => b.localeCompare(a));
  }, [transacoes]);

  const categoriasFiltro = useMemo(() => {
    return opcoesFiltroPorTipo(tipoFiltro, categorias, transacoes, 'categoria');
  }, [categorias, transacoes, tipoFiltro]);

  const pagamentosFiltro = useMemo(() => {
    return opcoesFiltroPorTipo(tipoFiltro, pagamentos, transacoes, 'formaPagamento');
  }, [pagamentos, transacoes, tipoFiltro]);

  const lancamentosFiltrados = useMemo(() => {
    return transacoes.filter((item) => {
      const passaMes = mesFiltro === 'todos' || item.data?.startsWith(mesFiltro);
      const passaPessoa = pessoaFiltro === 'todos' || item.pessoa === pessoaFiltro;
      const passaTipo = tipoFiltro === 'todos' || tipoDaTransacao(item) === tipoFiltro;
      const passaCategoria = categoriaFiltro === 'todos' || item.categoria === categoriaFiltro;
      const passaPagamento = pagamentoFiltro === 'todos' || item.formaPagamento === pagamentoFiltro;
      return passaMes && passaPessoa && passaTipo && passaCategoria && passaPagamento;
    });
  }, [transacoes, mesFiltro, pessoaFiltro, tipoFiltro, categoriaFiltro, pagamentoFiltro]);

  const resumoFiltrado = useMemo(() => calcularResumo(lancamentosFiltrados), [lancamentosFiltrados]);

  const historico = useMemo(() => {
    return [...lancamentosFiltrados].sort((a, b) => b.registradoEm - a.registradoEm);
  }, [lancamentosFiltrados]);

  const despesasPorCategoria = useMemo(() => {
    const totais = lancamentosFiltrados.reduce((acc, item) => {
      if (item.valor >= 0) return acc;
      const categoriaItem = labelsCategorias[item.categoria] || item.categoria || 'Outros';
      acc[categoriaItem] = (acc[categoriaItem] || 0) + Math.abs(item.valor);
      return acc;
    }, {});

    return Object.entries(totais)
      .map(([categoriaItem, total]) => ({ categoria: categoriaItem, total }))
      .sort((a, b) => b.total - a.total);
  }, [lancamentosFiltrados]);

  const maiorDespesaCategoria = useMemo(() => {
    return despesasPorCategoria.reduce((maior, item) => Math.max(maior, item.total), 0);
  }, [despesasPorCategoria]);

  const entradasPorCategoria = useMemo(() => {
    const totais = lancamentosFiltrados.reduce((acc, item) => {
      if (item.valor < 0) return acc;
      const categoriaItem = labelsCategorias[item.categoria] || item.categoria || 'Outros';
      acc[categoriaItem] = (acc[categoriaItem] || 0) + item.valor;
      return acc;
    }, {});

    return Object.entries(totais)
      .map(([categoriaItem, total]) => ({ categoria: categoriaItem, total }))
      .sort((a, b) => b.total - a.total);
  }, [lancamentosFiltrados]);

  const maiorEntradaCategoria = useMemo(() => {
    return entradasPorCategoria.reduce((maior, item) => Math.max(maior, item.total), 0);
  }, [entradasPorCategoria]);

  function calcularResumo(lista) {
    return lista.reduce(
      (acc, item) => {
        acc.saldo += item.valor;
        if (item.valor >= 0) acc.totalEntradas += item.valor;
        else acc.totalSaidas += Math.abs(item.valor);
        return acc;
      },
      { totalEntradas: 0, totalSaidas: 0, saldo: 0 },
    );
  }

  function gerarOcorrenciasRecorrentes(lista) {
    const limitePadrao = fimDoMesAtualISO();
    const ocorrencias = [];

    lista.forEach((regra) => {
      const intervalo = Math.max(1, Number(regra.intervalo) || 1);
      const maxOcorrencias = regra.vezes ? Math.max(1, Number(regra.vezes) || 1) : null;
      const limiteFinal = regra.dataFim && regra.dataFim < limitePadrao ? regra.dataFim : limitePadrao;

      if (!regra.dataInicio || regra.dataInicio > limiteFinal) return;

      let dataOcorrencia = dataLocal(regra.dataInicio);
      const diaReferencia = dataOcorrencia.getDate();
      let indice = 0;

      while (dataParaISO(dataOcorrencia) <= limiteFinal && indice < 600) {
        if (maxOcorrencias && indice >= maxOcorrencias) break;

        const dataISO = dataParaISO(dataOcorrencia);
        const valorBase = Number(regra.valor) || 0;

        ocorrencias.push({
          id: `recorrente-${regra.id}-${indice}`,
          recorrenteId: regra.id,
          recorrente: true,
          tipo: regra.tipo,
          data: dataISO,
          pessoa: regra.pessoa,
          categoria: regra.categoria,
          descricao: regra.descricao,
          formaPagamento: regra.formaPagamento,
          valor: regra.tipo === 'Saida' ? valorBase * -1 : valorBase,
          registradoEm: dataLocal(dataISO).getTime(),
        });

        indice += 1;
        dataOcorrencia = proximaDataRecorrente(dataOcorrencia, intervalo, regra.unidade, diaReferencia);
      }
    });

    return ocorrencias;
  }

  function trocarTipo(novoTipo) {
    setTipo(novoTipo);
    setCategoria('');
    setFormaPagamento('');
    setNovaCategoria('');
    setNovaFormaPagamento('');
  }

  function trocarTipoFiltro(novoTipo) {
    setTipoFiltro(novoTipo);

    const categoriasDoTipo = opcoesFiltroPorTipo(novoTipo, categorias, transacoes, 'categoria');
    const pagamentosDoTipo = opcoesFiltroPorTipo(novoTipo, pagamentos, transacoes, 'formaPagamento');

    setCategoriaFiltro((valorAtual) =>
      valorAtual === 'todos' || categoriasDoTipo.includes(valorAtual) ? valorAtual : 'todos',
    );
    setPagamentoFiltro((valorAtual) =>
      valorAtual === 'todos' || pagamentosDoTipo.includes(valorAtual) ? valorAtual : 'todos',
    );
  }

  function trocarTipoRecorrente(novoTipo) {
    setTipoRecorrente(novoTipo);
    setCategoriaRecorrente('');
    setFormaPagamentoRecorrente('');
  }

  function limparRecorrente() {
    setTipoRecorrente('Saida');
    setDataInicioRecorrente(hojeISO());
    setDataFimRecorrente('');
    setPessoaRecorrente('');
    setCategoriaRecorrente('');
    setDescricaoRecorrente('');
    setFormaPagamentoRecorrente('');
    setValorRecorrente('');
    setIntervaloRecorrente('1');
    setUnidadeRecorrente('meses');
    setVezesRecorrente('');
  }

  function cadastrarRecorrente(event) {
    event.preventDefault();
    setAlerta(null);

    if (
      !dataInicioRecorrente ||
      !pessoaRecorrente ||
      !categoriaRecorrente ||
      !formaPagamentoRecorrente ||
      !valorRecorrente ||
      !intervaloRecorrente
    ) {
      setAlerta({ tipo: 'erro', texto: 'Preencha inicio, nome, categoria, pagamento, valor e intervalo.' });
      return;
    }

    const valorNumero = normalizarValor(valorRecorrente);
    const intervaloNumero = Math.max(1, Number(intervaloRecorrente) || 1);
    const vezesNumero = vezesRecorrente ? Math.max(1, Number(vezesRecorrente) || 1) : '';

    if (valorNumero <= 0) {
      setAlerta({ tipo: 'erro', texto: 'Informe um valor maior que zero.' });
      return;
    }

    if (dataFimRecorrente && dataFimRecorrente < dataInicioRecorrente) {
      setAlerta({ tipo: 'erro', texto: 'A data de encerramento precisa ser depois do inicio.' });
      return;
    }

    const novaRegra = {
      id: crypto.randomUUID(),
      tipo: tipoRecorrente,
      dataInicio: dataInicioRecorrente,
      dataFim: dataFimRecorrente,
      pessoa: pessoaRecorrente,
      categoria: categoriaRecorrente,
      descricao: descricaoRecorrente.trim(),
      formaPagamento: formaPagamentoRecorrente,
      valor: valorNumero,
      intervalo: intervaloNumero,
      unidade: unidadeRecorrente,
      vezes: vezesNumero,
      criadoEm: Date.now(),
    };

    const proximasRegras = [...recorrentes, novaRegra];
    setRecorrentes(proximasRegras);
    salvarRecorrentes(proximasRegras);
    limparRecorrente();
    setAlerta({ tipo: 'sucesso', texto: 'Recorrencia cadastrada.' });
  }

  function excluirRecorrente(id) {
    const proximasRegras = recorrentes.filter((item) => item.id !== id);
    setRecorrentes(proximasRegras);
    salvarRecorrentes(proximasRegras);
    setAlerta({ tipo: 'sucesso', texto: 'Recorrencia removida.' });
  }

  function cadastrarPessoa() {
    const nome = novoNome.trim();
    if (!nome) return;

    const pessoaExistente = pessoas.find((item) => item.toLocaleLowerCase('pt-BR') === nome.toLocaleLowerCase('pt-BR'));
    if (pessoaExistente) {
      setPessoa(pessoaExistente);
      setNovoNome('');
      setAlerta({ tipo: 'sucesso', texto: 'Nome selecionado.' });
      return;
    }

    const proximasPessoas = [...pessoas, nome].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    setPessoas(proximasPessoas);
    salvarPessoas(proximasPessoas);
    setPessoa(nome);
    setNovoNome('');
    setAlerta({ tipo: 'sucesso', texto: 'Nome cadastrado.' });
  }

  function cadastrarOpcao({ valorNovo, lista, setLista, storageKey, setValorNovo, setValorSelecionado, mensagem }) {
    const nome = valorNovo.trim();
    if (!nome) return;

    const opcoesDoTipo = lista[tipo] || [];
    const opcaoExistente = opcoesDoTipo.find((item) => item.toLocaleLowerCase('pt-BR') === nome.toLocaleLowerCase('pt-BR'));
    if (opcaoExistente) {
      setValorSelecionado(opcaoExistente);
      setValorNovo('');
      setAlerta({ tipo: 'sucesso', texto: 'Opcao selecionada.' });
      return;
    }

    const proximasOpcoes = {
      ...lista,
      [tipo]: [...opcoesDoTipo, nome].sort((a, b) => a.localeCompare(b, 'pt-BR')),
    };
    setLista(proximasOpcoes);
    salvarOpcoes(storageKey, proximasOpcoes);
    setValorSelecionado(nome);
    setValorNovo('');
    setAlerta({ tipo: 'sucesso', texto: mensagem });
  }

  function excluirOpcao({ item, lista, setLista, storageKey, setValorSelecionado }) {
    const proximasOpcoes = {
      ...lista,
      [tipo]: (lista[tipo] || []).filter((opcao) => opcao !== item),
    };

    setLista(proximasOpcoes);
    salvarOpcoes(storageKey, proximasOpcoes);
    setValorSelecionado((valorAtual) => (valorAtual === item ? '' : valorAtual));
    if (storageKey === CATEGORIAS_STORAGE_KEY) {
      setCategoriaFiltro((valorAtual) => (valorAtual === item ? 'todos' : valorAtual));
    }
    if (storageKey === PAGAMENTOS_STORAGE_KEY) {
      setPagamentoFiltro((valorAtual) => (valorAtual === item ? 'todos' : valorAtual));
    }
    setAlerta({ tipo: 'sucesso', texto: 'Cadastro removido da lista.' });
  }

  function excluirPessoa(nome) {
    const proximasPessoas = pessoas.filter((item) => item !== nome);
    setPessoas(proximasPessoas);
    salvarPessoas(proximasPessoas);
    setPessoa((valorAtual) => (valorAtual === nome ? '' : valorAtual));
    setPessoaFiltro((valorAtual) => (valorAtual === nome ? 'todos' : valorAtual));
    setAlerta({ tipo: 'sucesso', texto: 'Nome removido da lista.' });
  }

  function registrarLancamento(event) {
    event.preventDefault();
    setAlerta(null);

    if (!data || !pessoa || !categoria || !formaPagamento || !valor) {
      setAlerta({ tipo: 'erro', texto: 'Preencha nome, data, categoria, forma de pagamento e valor.' });
      return;
    }

    const valorNumero = normalizarValor(valor);
    if (valorNumero <= 0) {
      setAlerta({ tipo: 'erro', texto: 'Informe um valor maior que zero.' });
      return;
    }

    const novoLancamento = {
      id: crypto.randomUUID(),
      tipo,
      data,
      pessoa,
      categoria,
      descricao: descricao.trim(),
      formaPagamento,
      valor: tipo === 'Saida' ? valorNumero * -1 : valorNumero,
      registradoEm: Date.now(),
    };

    const proximosLancamentos = [...lancamentos, novoLancamento];
    setLancamentos(proximosLancamentos);
    salvarLancamentos(proximosLancamentos);

    setTipo('Saida');
    setPessoa('');
    setCategoria('');
    setNovaCategoria('');
    setDescricao('');
    setFormaPagamento('');
    setNovaFormaPagamento('');
    setValor('');
    setData(hojeISO());
    setAlerta({ tipo: 'sucesso', texto: 'Lancamento registrado com sucesso.' });
  }

  function removerLancamento(id) {
    const proximosLancamentos = lancamentos.filter((item) => item.id !== id);
    setLancamentos(proximosLancamentos);
    salvarLancamentos(proximosLancamentos);
  }

  function exportarExcel() {
    const linhas = historico.map((item) => ({
      Data: formatarData(item.data),
      Pessoa: item.pessoa || 'Sem pessoa',
      Categoria: labelsCategorias[item.categoria] || item.categoria,
      Descricao: item.descricao,
      'Forma de Pagamento': labelsPagamentos[item.formaPagamento] || item.formaPagamento,
      Valor: item.valor,
      Recorrente: item.recorrente ? 'Sim' : 'Nao',
      'Data de Registro': new Date(item.registradoEm).toLocaleString('pt-BR'),
    }));

    const worksheet = XLSX.utils.json_to_sheet(linhas);
    worksheet['!cols'] = [
      { wch: 12 },
      { wch: 18 },
      { wch: 18 },
      { wch: 28 },
      { wch: 22 },
      { wch: 14 },
      { wch: 12 },
      { wch: 22 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Extrato');

    const mesArquivo = mesFiltro === 'todos' ? 'todos-os-meses' : mesFiltro;
    const pessoaArquivo = pessoaFiltro === 'todos' ? 'todas-as-pessoas' : pessoaFiltro.replace(/\s+/g, '-').toLowerCase();
    const tipoArquivo = tipoFiltro === 'todos' ? 'entradas-e-saidas' : tipoFiltro.toLowerCase();
    const categoriaArquivo =
      categoriaFiltro === 'todos' ? 'todas-as-categorias' : categoriaFiltro.replace(/\s+/g, '-').toLowerCase();
    const pagamentoArquivo =
      pagamentoFiltro === 'todos' ? 'todas-as-formas' : pagamentoFiltro.replace(/\s+/g, '-').toLowerCase();
    XLSX.writeFile(
      workbook,
      `controle-de-gastos-${mesArquivo}-${pessoaArquivo}-${tipoArquivo}-${categoriaArquivo}-${pagamentoArquivo}.xlsx`,
    );
  }

  return (
    <main className="page-shell">
      <section className="app-shell" aria-label="Controle de Gastos 2.0">
        <header className="topbar">
          <div>
            <p className="eyebrow">Controle de Gastos 2.0</p>
            <h1>Meus lancamentos</h1>
          </div>
          <div className="saldo-pill">
            <Wallet size={18} />
            <span>{formatarMoeda(resumoGeral.saldo)}</span>
          </div>
        </header>

        <nav className="tabs-nav" aria-label="Abas principais">
          <button className={abaAtiva === 'form' ? 'tab-btn active' : 'tab-btn'} onClick={() => setAbaAtiva('form')}>
            <Plus size={18} />
            <span>Novo</span>
          </button>
          <button
            className={abaAtiva === 'extrato' ? 'tab-btn active' : 'tab-btn'}
            onClick={() => setAbaAtiva('extrato')}
          >
            <CalendarDays size={18} />
            <span>Extrato</span>
          </button>
          <button
            className={abaAtiva === 'recorrentes' ? 'tab-btn active' : 'tab-btn'}
            onClick={() => setAbaAtiva('recorrentes')}
          >
            <Repeat2 size={18} />
            <span>Recorrentes</span>
          </button>
          <button
            className={abaAtiva === 'grafico' ? 'tab-btn active' : 'tab-btn'}
            onClick={() => setAbaAtiva('grafico')}
          >
            <BarChart3 size={18} />
            <span>Grafico</span>
          </button>
          <button
            className={abaAtiva === 'cadastros' ? 'tab-btn active' : 'tab-btn'}
            onClick={() => setAbaAtiva('cadastros')}
          >
            <UserPlus size={18} />
            <span>Cadastros</span>
          </button>
        </nav>

        {abaAtiva === 'form' ? (
          <form className="panel" onSubmit={registrarLancamento}>
            <div className="panel-heading">
              <h2>Novo lancamento</h2>
              <div className={tipo === 'Entrada' ? 'type-badge entrada' : 'type-badge saida'}>
                {tipo === 'Entrada' ? <ArrowUpCircle size={16} /> : <ArrowDownCircle size={16} />}
                <span>{tipo === 'Entrada' ? 'Entrada' : 'Saida'}</span>
              </div>
            </div>

            {alerta && <div className={`alert ${alerta.tipo}`}>{alerta.texto}</div>}

            <div className="field-group">
              <label htmlFor="tipo">Tipo de lancamento *</label>
              <select id="tipo" value={tipo} onChange={(event) => trocarTipo(event.target.value)}>
                <option value="Saida">Saida (Despesa)</option>
                <option value="Entrada">Entrada (Receita)</option>
              </select>
            </div>

            <div className="field-grid">
              <div className="field-group">
                <label htmlFor="data">Data *</label>
                <input id="data" type="date" value={data} onChange={(event) => setData(event.target.value)} />
              </div>
              <div className="field-group">
                <label htmlFor="valor">Valor (R$) *</label>
                <input
                  id="valor"
                  inputMode="decimal"
                  placeholder="0,00"
                  value={valor}
                  onChange={(event) => setValor(event.target.value)}
                />
              </div>
            </div>

            <div className="field-group">
              <label htmlFor="categoria">Categoria / Fonte *</label>
              <select id="categoria" value={categoria} onChange={(event) => setCategoria(event.target.value)}>
                <option value="">Selecione...</option>
                {(categorias[tipo] || []).map((item) => (
                  <option key={item} value={item}>
                    {labelsCategorias[item] || item}
                  </option>
                ))}
              </select>
            </div>

            <div className="field-group">
              <label htmlFor="descricao">Descricao</label>
              <input
                id="descricao"
                placeholder="Ex: salario mensal, supermercado..."
                value={descricao}
                onChange={(event) => setDescricao(event.target.value)}
              />
            </div>

            <div className="field-group">
              <label htmlFor="formaPagamento">Forma de pagamento *</label>
              <select
                id="formaPagamento"
                value={formaPagamento}
                onChange={(event) => setFormaPagamento(event.target.value)}
              >
                <option value="">Selecione...</option>
                {(pagamentos[tipo] || []).map((item) => (
                  <option key={item} value={item}>
                    {labelsPagamentos[item] || item}
                  </option>
                ))}
              </select>
            </div>

            <div className="field-group">
              <label htmlFor="pessoa">Nome *</label>
              <select id="pessoa" value={pessoa} onChange={(event) => setPessoa(event.target.value)}>
                <option value="">Selecione...</option>
                {pessoas.map((nome) => (
                  <option key={nome} value={nome}>
                    {nome}
                  </option>
                ))}
              </select>
            </div>

            <button className="primary-btn" type="submit">
              <Plus size={18} />
              <span>Registrar lancamento</span>
            </button>
          </form>
        ) : abaAtiva === 'recorrentes' ? (
          <section className="panel">
            <div className="panel-heading">
              <h2>Recorrentes</h2>
              <div className={tipoRecorrente === 'Entrada' ? 'type-badge entrada' : 'type-badge saida'}>
                {tipoRecorrente === 'Entrada' ? <ArrowUpCircle size={16} /> : <ArrowDownCircle size={16} />}
                <span>{tipoRecorrente === 'Entrada' ? 'Entrada' : 'Saida'}</span>
              </div>
            </div>

            {alerta && <div className={`alert ${alerta.tipo}`}>{alerta.texto}</div>}

            <form className="recurring-form" onSubmit={cadastrarRecorrente}>
              <div className="field-group">
                <label htmlFor="tipoRecorrente">Tipo *</label>
                <select id="tipoRecorrente" value={tipoRecorrente} onChange={(event) => trocarTipoRecorrente(event.target.value)}>
                  <option value="Saida">Saida (Despesa)</option>
                  <option value="Entrada">Entrada (Receita)</option>
                </select>
              </div>

              <div className="field-grid">
                <div className="field-group">
                  <label htmlFor="dataInicioRecorrente">Comeca em *</label>
                  <input
                    id="dataInicioRecorrente"
                    type="date"
                    value={dataInicioRecorrente}
                    onChange={(event) => setDataInicioRecorrente(event.target.value)}
                  />
                </div>
                <div className="field-group">
                  <label htmlFor="valorRecorrente">Valor (R$) *</label>
                  <input
                    id="valorRecorrente"
                    inputMode="decimal"
                    placeholder="0,00"
                    value={valorRecorrente}
                    onChange={(event) => setValorRecorrente(event.target.value)}
                  />
                </div>
              </div>

              <div className="field-grid">
                <div className="field-group">
                  <label htmlFor="categoriaRecorrente">Categoria / Fonte *</label>
                  <select
                    id="categoriaRecorrente"
                    value={categoriaRecorrente}
                    onChange={(event) => setCategoriaRecorrente(event.target.value)}
                  >
                    <option value="">Selecione...</option>
                    {(categorias[tipoRecorrente] || []).map((item) => (
                      <option key={item} value={item}>
                        {labelsCategorias[item] || item}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="field-group">
                  <label htmlFor="formaPagamentoRecorrente">Forma de pagamento *</label>
                  <select
                    id="formaPagamentoRecorrente"
                    value={formaPagamentoRecorrente}
                    onChange={(event) => setFormaPagamentoRecorrente(event.target.value)}
                  >
                    <option value="">Selecione...</option>
                    {(pagamentos[tipoRecorrente] || []).map((item) => (
                      <option key={item} value={item}>
                        {labelsPagamentos[item] || item}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="field-group">
                <label htmlFor="pessoaRecorrente">Nome *</label>
                <select id="pessoaRecorrente" value={pessoaRecorrente} onChange={(event) => setPessoaRecorrente(event.target.value)}>
                  <option value="">Selecione...</option>
                  {pessoas.map((nome) => (
                    <option key={nome} value={nome}>
                      {nome}
                    </option>
                  ))}
                </select>
              </div>

              <div className="field-group">
                <label htmlFor="descricaoRecorrente">Descricao</label>
                <input
                  id="descricaoRecorrente"
                  placeholder="Ex: aluguel, salario mensal..."
                  value={descricaoRecorrente}
                  onChange={(event) => setDescricaoRecorrente(event.target.value)}
                />
              </div>

              <div className="recurrence-row">
                <div className="field-group">
                  <label htmlFor="intervaloRecorrente">Repetir a cada *</label>
                  <input
                    id="intervaloRecorrente"
                    type="number"
                    min="1"
                    value={intervaloRecorrente}
                    onChange={(event) => setIntervaloRecorrente(event.target.value)}
                  />
                </div>
                <div className="field-group">
                  <label htmlFor="unidadeRecorrente">Periodo *</label>
                  <select id="unidadeRecorrente" value={unidadeRecorrente} onChange={(event) => setUnidadeRecorrente(event.target.value)}>
                    <option value="dias">Dias</option>
                    <option value="semanas">Semanas</option>
                    <option value="meses">Meses</option>
                    <option value="anos">Anos</option>
                  </select>
                </div>
              </div>

              <div className="field-grid">
                <div className="field-group">
                  <label htmlFor="vezesRecorrente">Quantidade de vezes</label>
                  <input
                    id="vezesRecorrente"
                    type="number"
                    min="1"
                    placeholder="Opcional"
                    value={vezesRecorrente}
                    onChange={(event) => setVezesRecorrente(event.target.value)}
                  />
                </div>
                <div className="field-group">
                  <label htmlFor="dataFimRecorrente">Encerra em</label>
                  <input
                    id="dataFimRecorrente"
                    type="date"
                    value={dataFimRecorrente}
                    onChange={(event) => setDataFimRecorrente(event.target.value)}
                  />
                </div>
              </div>

              <button className="primary-btn" type="submit">
                <Repeat2 size={18} />
                <span>Cadastrar recorrencia</span>
              </button>
            </form>

            <div className="recurring-list">
              {recorrentes.length ? (
                recorrentes.map((item) => (
                  <article className={item.tipo === 'Entrada' ? 'recurring-item entrada' : 'recurring-item saida'} key={item.id}>
                    <div>
                      <div className="item-categoria">{labelsCategorias[item.categoria] || item.categoria}</div>
                      <div className="item-desc">{item.descricao || 'Sem descricao'}</div>
                      <div className="item-meta">
                        {item.pessoa} - {labelsPagamentos[item.formaPagamento] || item.formaPagamento}
                      </div>
                      <div className="item-meta">
                        Desde {formatarData(item.dataInicio)} - a cada {item.intervalo}{' '}
                        {item.unidade}
                        {item.vezes ? ` - ${item.vezes} vezes` : ''}
                        {item.dataFim ? ` - ate ${formatarData(item.dataFim)}` : ''}
                      </div>
                    </div>
                    <div className="item-actions">
                      <strong className={item.tipo === 'Entrada' ? 'valor-positivo' : 'valor-negativo'}>
                        {formatarMoeda(item.valor)}
                      </strong>
                      <button className="delete-btn" onClick={() => excluirRecorrente(item.id)} title="Remover recorrencia">
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </article>
                ))
              ) : (
                <div className="estado-vazio">Nenhuma recorrencia cadastrada.</div>
              )}
            </div>
          </section>
        ) : abaAtiva === 'cadastros' ? (
          <section className="panel">
            <div className="panel-heading">
              <h2>Cadastros</h2>
              <div className={tipo === 'Entrada' ? 'type-badge entrada' : 'type-badge saida'}>
                {tipo === 'Entrada' ? <ArrowUpCircle size={16} /> : <ArrowDownCircle size={16} />}
                <span>{tipo === 'Entrada' ? 'Entrada' : 'Saida'}</span>
              </div>
            </div>

            {alerta && <div className={`alert ${alerta.tipo}`}>{alerta.texto}</div>}

            <div className="field-group">
              <label htmlFor="tipoCadastro">Tipo para categoria e pagamento</label>
              <select id="tipoCadastro" value={tipo} onChange={(event) => trocarTipo(event.target.value)}>
                <option value="Saida">Saida (Despesa)</option>
                <option value="Entrada">Entrada (Receita)</option>
              </select>
            </div>

            <CadastroBloco
              titulo={tipo === 'Entrada' ? 'Fonte de entrada' : 'Categoria de despesa'}
              lista={categorias[tipo] || []}
              valor={novaCategoria}
              placeholder={tipo === 'Entrada' ? 'Nova fonte' : 'Nova categoria'}
              onChange={setNovaCategoria}
              onCadastrar={() =>
                cadastrarOpcao({
                  valorNovo: novaCategoria,
                  lista: categorias,
                  setLista: setCategorias,
                  storageKey: CATEGORIAS_STORAGE_KEY,
                  setValorNovo: setNovaCategoria,
                  setValorSelecionado: setCategoria,
                  mensagem: tipo === 'Entrada' ? 'Fonte cadastrada.' : 'Categoria cadastrada.',
                })
              }
              onExcluir={(item) =>
                excluirOpcao({
                  item,
                  lista: categorias,
                  setLista: setCategorias,
                  storageKey: CATEGORIAS_STORAGE_KEY,
                  setValorSelecionado: setCategoria,
                })
              }
            />

            <CadastroBloco
              titulo="Forma de pagamento"
              lista={pagamentos[tipo] || []}
              valor={novaFormaPagamento}
              placeholder="Nova forma de pagamento"
              onChange={setNovaFormaPagamento}
              onCadastrar={() =>
                cadastrarOpcao({
                  valorNovo: novaFormaPagamento,
                  lista: pagamentos,
                  setLista: setPagamentos,
                  storageKey: PAGAMENTOS_STORAGE_KEY,
                  setValorNovo: setNovaFormaPagamento,
                  setValorSelecionado: setFormaPagamento,
                  mensagem: 'Forma de pagamento cadastrada.',
                })
              }
              onExcluir={(item) =>
                excluirOpcao({
                  item,
                  lista: pagamentos,
                  setLista: setPagamentos,
                  storageKey: PAGAMENTOS_STORAGE_KEY,
                  setValorSelecionado: setFormaPagamento,
                })
              }
            />

            <CadastroBloco
              titulo="Nomes"
              lista={pessoas}
              valor={novoNome}
              placeholder="Novo nome"
              onChange={setNovoNome}
              onCadastrar={cadastrarPessoa}
              onExcluir={excluirPessoa}
            />
          </section>
        ) : abaAtiva === 'extrato' ? (
          <section className="panel">
            <div className="panel-heading">
              <h2>Resumo financeiro</h2>
              <button className="icon-btn" onClick={exportarExcel} title="Baixar Excel" disabled={!historico.length}>
                <Download size={18} />
              </button>
            </div>

            <div className="filter-row">
              <div className="field-group filter-field">
                <label htmlFor="mesFiltro">Mes do extrato</label>
                <select id="mesFiltro" value={mesFiltro} onChange={(event) => setMesFiltro(event.target.value)}>
                  <option value="todos">Todos os meses</option>
                  {mesesDisponiveis.map((mes) => (
                    <option key={mes} value={mes}>
                      {formatarMes(mes)}
                    </option>
                  ))}
                </select>
              </div>

              <div className="field-group filter-field">
                <label htmlFor="pessoaFiltro">Pessoa</label>
                <select id="pessoaFiltro" value={pessoaFiltro} onChange={(event) => setPessoaFiltro(event.target.value)}>
                  <option value="todos">Todas as pessoas</option>
                  {pessoas.map((nome) => (
                    <option key={nome} value={nome}>
                      {nome}
                    </option>
                  ))}
                </select>
              </div>

              <div className="field-group filter-field">
                <label htmlFor="tipoFiltro">Tipo</label>
                <select id="tipoFiltro" value={tipoFiltro} onChange={(event) => trocarTipoFiltro(event.target.value)}>
                  <option value="todos">Entrada e saida</option>
                  <option value="Entrada">Entrada</option>
                  <option value="Saida">Saida</option>
                </select>
              </div>

              <div className="field-group filter-field">
                <label htmlFor="categoriaFiltro">Categoria / Fonte</label>
                <select
                  id="categoriaFiltro"
                  value={categoriaFiltro}
                  onChange={(event) => setCategoriaFiltro(event.target.value)}
                >
                  <option value="todos">Todas as categorias/fontes</option>
                  {categoriasFiltro.map((item) => (
                    <option key={item} value={item}>
                      {labelsCategorias[item] || item}
                    </option>
                  ))}
                </select>
              </div>

              <div className="field-group filter-field">
                <label htmlFor="pagamentoFiltro">Forma de pagamento</label>
                <select
                  id="pagamentoFiltro"
                  value={pagamentoFiltro}
                  onChange={(event) => setPagamentoFiltro(event.target.value)}
                >
                  <option value="todos">Todas as formas</option>
                  {pagamentosFiltro.map((item) => (
                    <option key={item} value={item}>
                      {labelsPagamentos[item] || item}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            <div className="dashboard">
              <ResumoCard label="Entradas" valor={resumoFiltrado.totalEntradas} status="positivo" />
              <ResumoCard label="Saidas" valor={resumoFiltrado.totalSaidas} status="negativo" />
              <ResumoCard
                label="Saldo"
                valor={resumoFiltrado.saldo}
                status={resumoFiltrado.saldo >= 0 ? 'positivo' : 'negativo'}
                destaque
              />
            </div>

            <div className="extrato-header">
              <h3>Historico</h3>
              <button className="ghost-btn" onClick={() => setLancamentos(carregarLancamentos())}>
                <RefreshCw size={16} />
                <span>Atualizar</span>
              </button>
            </div>

            <div className="lista-container">
              {historico.length ? (
                historico.map((item) => (
                  <article
                    className={item.valor >= 0 ? 'extrato-item cartao-entrada' : 'extrato-item cartao-saida'}
                    key={item.id}
                  >
                    <div className="item-info">
                      <div className="item-categoria">{labelsCategorias[item.categoria] || item.categoria}</div>
                      <div className="item-desc">{item.descricao || 'Sem descricao'}</div>
                      <div className="item-meta">
                        {item.pessoa || 'Sem pessoa'} - {formatarData(item.data)} -{' '}
                        {labelsPagamentos[item.formaPagamento] || item.formaPagamento}
                      </div>
                      {item.recorrente && <div className="item-meta recorrente-meta">Recorrente</div>}
                    </div>
                    <div className="item-actions">
                      <strong className={item.valor >= 0 ? 'valor-positivo' : 'valor-negativo'}>
                        {formatarMoeda(item.valor)}
                      </strong>
                      {item.recorrente ? (
                        <span className="recurring-lock">Auto</span>
                      ) : (
                        <button className="delete-btn" onClick={() => removerLancamento(item.id)} title="Remover">
                          <Trash2 size={16} />
                        </button>
                      )}
                    </div>
                  </article>
                ))
              ) : (
                <div className="estado-vazio">
                  Nenhum lancamento para os filtros selecionados.
                </div>
              )}
            </div>
          </section>
        ) : (
          <section className="panel">
            <div className="panel-heading">
              <h2>Despesas por categoria</h2>
              <div className="type-badge saida">
                <BarChart3 size={16} />
                <span>{formatarMoeda(resumoFiltrado.totalSaidas)}</span>
              </div>
            </div>

            <div className="filter-row">
              <div className="field-group filter-field">
                <label htmlFor="mesFiltroGrafico">Mes do grafico</label>
                <select id="mesFiltroGrafico" value={mesFiltro} onChange={(event) => setMesFiltro(event.target.value)}>
                  <option value="todos">Todos os meses</option>
                  {mesesDisponiveis.map((mes) => (
                    <option key={mes} value={mes}>
                      {formatarMes(mes)}
                    </option>
                  ))}
                </select>
              </div>

              <div className="field-group filter-field">
                <label htmlFor="pessoaFiltroGrafico">Pessoa</label>
                <select
                  id="pessoaFiltroGrafico"
                  value={pessoaFiltro}
                  onChange={(event) => setPessoaFiltro(event.target.value)}
                >
                  <option value="todos">Todas as pessoas</option>
                  {pessoas.map((nome) => (
                    <option key={nome} value={nome}>
                      {nome}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {despesasPorCategoria.length ? (
              <div className="chart-panel" aria-label="Grafico de despesas por categoria">
                <div className="column-chart">
                  {despesasPorCategoria.map((item) => {
                    const altura = maiorDespesaCategoria ? Math.max((item.total / maiorDespesaCategoria) * 100, 8) : 0;
                    return (
                      <div className="chart-column" key={item.categoria}>
                        <div className="bar-value">{formatarMoeda(item.total)}</div>
                        <div className="bar-track">
                          <div className="bar-fill" style={{ height: `${altura}%` }} />
                        </div>
                        <div className="bar-label">{item.categoria}</div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (
              <div className="estado-vazio">Nenhuma despesa para montar o grafico.</div>
            )}

            <div className="chart-section-heading">
              <h2>Entradas por categoria</h2>
              <div className="type-badge entrada">
                <BarChart3 size={16} />
                <span>{formatarMoeda(resumoFiltrado.totalEntradas)}</span>
              </div>
            </div>

            {entradasPorCategoria.length ? (
              <div className="chart-panel" aria-label="Grafico de entradas por categoria">
                <div className="column-chart">
                  {entradasPorCategoria.map((item) => {
                    const altura = maiorEntradaCategoria ? Math.max((item.total / maiorEntradaCategoria) * 100, 8) : 0;
                    return (
                      <div className="chart-column" key={item.categoria}>
                        <div className="bar-value entrada">{formatarMoeda(item.total)}</div>
                        <div className="bar-track">
                          <div className="bar-fill entrada" style={{ height: `${altura}%` }} />
                        </div>
                        <div className="bar-label">{item.categoria}</div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (
              <div className="estado-vazio">Nenhuma entrada para montar o grafico.</div>
            )}
          </section>
        )}
      </section>
    </main>
  );
}

function ResumoCard({ label, valor, status, destaque = false }) {
  return (
    <div className={destaque ? 'dash-card dash-saldo' : 'dash-card'}>
      <small>{label}</small>
      <strong className={status === 'positivo' ? 'valor-positivo' : 'valor-negativo'}>{formatarMoeda(valor)}</strong>
    </div>
  );
}

function CadastroBloco({ titulo, lista, valor, placeholder, onChange, onCadastrar, onExcluir }) {
  return (
    <section className="register-section">
      <div className="register-heading">
        <h3>{titulo}</h3>
        <span>{lista.length} itens</span>
      </div>

      <div className="add-person-row">
        <input
          aria-label={placeholder}
          placeholder={placeholder}
          value={valor}
          onChange={(event) => onChange(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              onCadastrar();
            }
          }}
        />
        <button className="small-action-btn" type="button" onClick={onCadastrar}>
          <Plus size={18} />
          <span>Cadastrar</span>
        </button>
      </div>

      <div className="tag-list">
        {lista.map((item) => (
          <div className="tag-item" key={item}>
            <span>{labelsCategorias[item] || labelsPagamentos[item] || item}</span>
            <button type="button" onClick={() => onExcluir(item)} title="Excluir">
              <Trash2 size={13} />
            </button>
          </div>
        ))}
      </div>
    </section>
  );
}

createRoot(document.getElementById('root')).render(<App />);
