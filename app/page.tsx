"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import { supabase } from "../lib/supabase";

type Compra = {
  id: number;
  descricao: string;
  valor: number;
  cartao: string;
  pessoa: string;
  parcelas: number;
  parcelaAtual: number;
  data: string;
  tipo?: "parcela" | "avulso";
};

type Cartao = {
  nome: string;
  limite: number;
};

type FaturaCartao = {
  nome: string;
  fechamento: number;
  vencimento: number;
};

type Lancamento = {
  id: number;
  descricao: string;
  valor: number;
};

type Assinatura = {
  id: number;
  descricao: string;
  valor: number;
  cartao: string;
  pessoa: string;
};

type HistoricoMensal = {
  mes: string;
  receitas: number;
  despesasMensais: number;
  cartoes: number;
  assinaturas: number;
  despesasTotais: number;
  parcelasFuturas: number;
  sobra: number;
  criadoEm: string;
};

type AppData = {
  compras: Compra[];
  cartoes: Cartao[];
  faturasCartoes: FaturaCartao[];
  pessoas: string[];
  pessoasSelecionadasCartoes: string[];
  pessoasSelecionadasParcelas: string[];
  receitasClt: Lancamento[];
  receitasMei: Lancamento[];
  despesasMensais: Lancamento[];
  assinaturas: Assinatura[];
  historicoMensal: HistoricoMensal[];
  ultimaAtualizacaoMensal: string;
};

const APP_DATA_ID = "controle-financeiro";

const cartoesIniciais: Cartao[] = [
  { nome: "Platinum", limite: 5000 },
  { nome: "Uniclass", limite: 5000 },
  { nome: "Latam", limite: 5000 },
  { nome: "Mercado Pago", limite: 5000 },
];

const faturasCartoesIniciais: FaturaCartao[] = [
  { nome: "Platinum", fechamento: 15, vencimento: 22 },
  { nome: "Uniclass", fechamento: 15, vencimento: 22 },
  { nome: "Latam", fechamento: 10, vencimento: 17 },
  { nome: "Mercado Pago", fechamento: 5, vencimento: 12 },
];

const pessoasIniciais = [
  "Rennan",
  "Bruno",
  "Amaral",
  "Tio",
  "Elis",
  "Gloria",
  "Mey",
];

const receitasCltIniciais: Lancamento[] = [
  { id: 1, descricao: "Salário", valor: 2100 },
  { id: 2, descricao: "Extra", valor: 0 },
];

const receitasMeiIniciais: Lancamento[] = [
  { id: 1, descricao: "Pastor", valor: 160 },
  { id: 2, descricao: "Juan", valor: 1000 },
  { id: 3, descricao: "Janira", valor: 250 },
  { id: 4, descricao: "Pablo", valor: 350 },
  { id: 5, descricao: "Lidia", valor: 500 },
];

const despesasMensaisIniciais: Lancamento[] = [
  { id: 1, descricao: "Faculdade", valor: 349.68 },
  { id: 2, descricao: "Econtador", valor: 168.47 },
  { id: 3, descricao: "Wallace", valor: 500 },
  { id: 4, descricao: "Claro", valor: 227.02 },
  { id: 5, descricao: "Mãe", valor: 454.98 },
];

function comprasComTipo(lista: Compra[] = []) {
  return lista.map((item) => ({
    ...item,
    tipo: item.tipo || (item.parcelas > 1 ? "parcela" : "avulso"),
  }));
}

function assinaturasComPessoa(lista: Assinatura[] = []) {
  return lista.map((item) => ({
    ...item,
    pessoa: item.pessoa || "Rennan",
  }));
}

function dadosPadrao(): AppData {
  return {
    compras: [],
    cartoes: cartoesIniciais,
    faturasCartoes: faturasCartoesIniciais,
    pessoas: pessoasIniciais,
    pessoasSelecionadasCartoes: ["Rennan"],
    pessoasSelecionadasParcelas: ["Rennan"],
    receitasClt: receitasCltIniciais,
    receitasMei: receitasMeiIniciais,
    despesasMensais: despesasMensaisIniciais,
    assinaturas: [],
    historicoMensal: [],
    ultimaAtualizacaoMensal: chaveMesAtual(),
  };
}

function normalizarDadosOnline(data: Partial<AppData> | null | undefined): AppData {
  const padrao = dadosPadrao();

  if (!data || Object.keys(data).length === 0) {
    return padrao;
  }

  return {
    compras: comprasComTipo(data.compras || []),
    cartoes: data.cartoes?.length ? data.cartoes : padrao.cartoes,
    faturasCartoes: data.faturasCartoes?.length
      ? data.faturasCartoes
      : padrao.faturasCartoes,
    pessoas: data.pessoas?.length ? data.pessoas : padrao.pessoas,
    pessoasSelecionadasCartoes:
      data.pessoasSelecionadasCartoes?.length
        ? data.pessoasSelecionadasCartoes
        : padrao.pessoasSelecionadasCartoes,
    pessoasSelecionadasParcelas:
      data.pessoasSelecionadasParcelas?.length
        ? data.pessoasSelecionadasParcelas
        : padrao.pessoasSelecionadasParcelas,
    receitasClt: data.receitasClt?.length ? data.receitasClt : padrao.receitasClt,
    receitasMei: data.receitasMei?.length ? data.receitasMei : padrao.receitasMei,
    despesasMensais: data.despesasMensais?.length
      ? data.despesasMensais
      : padrao.despesasMensais,
    assinaturas: assinaturasComPessoa(data.assinaturas || []),
    historicoMensal: data.historicoMensal || [],
    ultimaAtualizacaoMensal: data.ultimaAtualizacaoMensal || chaveMesAtual(),
  };
}

function montarDadosDoLocalStorage(): AppData {
  const padrao = dadosPadrao();

  try {
    const receitasAntigas = localStorage.getItem("receitas");
    const receitasAntigasLista: Lancamento[] = receitasAntigas
      ? JSON.parse(receitasAntigas)
      : [];

    const receitasCltMigradas = receitasAntigasLista.filter((item) =>
      ["salário", "salario", "extra"].includes(item.descricao.toLowerCase()),
    );

    const receitasMeiMigradas = receitasAntigasLista.filter(
      (item) =>
        !["salário", "salario", "extra"].includes(item.descricao.toLowerCase()),
    );

    return normalizarDadosOnline({
      compras: JSON.parse(localStorage.getItem("compras") || "[]"),
      cartoes: JSON.parse(localStorage.getItem("cartoes") || "null") || padrao.cartoes,
      faturasCartoes:
        JSON.parse(localStorage.getItem("faturasCartoes") || "null") ||
        padrao.faturasCartoes,
      pessoas: JSON.parse(localStorage.getItem("pessoas") || "null") || padrao.pessoas,
      pessoasSelecionadasCartoes:
        JSON.parse(localStorage.getItem("pessoasSelecionadasCartoes") || "null") ||
        padrao.pessoasSelecionadasCartoes,
      pessoasSelecionadasParcelas:
        JSON.parse(localStorage.getItem("pessoasSelecionadasParcelas") || "null") ||
        padrao.pessoasSelecionadasParcelas,
      receitasClt:
        JSON.parse(localStorage.getItem("receitasClt") || "null") ||
        receitasCltMigradas ||
        padrao.receitasClt,
      receitasMei:
        JSON.parse(localStorage.getItem("receitasMei") || "null") ||
        receitasMeiMigradas ||
        padrao.receitasMei,
      despesasMensais:
        JSON.parse(localStorage.getItem("despesasMensais") || "null") ||
        padrao.despesasMensais,
      assinaturas: JSON.parse(localStorage.getItem("assinaturas") || "[]"),
      historicoMensal: JSON.parse(localStorage.getItem("historicoMensal") || "[]"),
      ultimaAtualizacaoMensal:
        localStorage.getItem("ultimaAtualizacaoMensal") || chaveMesAtual(),
    });
  } catch {
    return padrao;
  }
}

function calcularHistoricoPorDados(mes: string, dados: AppData): HistoricoMensal {
  const receitas =
    dados.receitasClt.reduce((soma, item) => soma + item.valor, 0) +
    dados.receitasMei.reduce((soma, item) => soma + item.valor, 0);

  const despesasMensais = dados.despesasMensais.reduce(
    (soma, item) => soma + item.valor,
    0,
  );

  const cartoes = dados.compras
    .filter((item) => item.pessoa === "Rennan")
    .reduce((soma, item) => soma + item.valor, 0);

  const assinaturas = dados.assinaturas
    .filter((item) => (item.pessoa || "Rennan") === "Rennan")
    .reduce((soma, item) => soma + item.valor, 0);

  const despesasTotais = despesasMensais + cartoes + assinaturas;

  const parcelasFuturas = dados.compras
    .filter((item) => item.pessoa === "Rennan")
    .reduce((soma, item) => {
      if (tipoDaCompra(item) !== "parcela") return soma;

      const atual = item.parcelaAtual || 1;
      const restantesIncluindoAtual = Math.max(item.parcelas - atual + 1, 0);

      return soma + item.valor * restantesIncluindoAtual;
    }, 0);

  return {
    mes,
    receitas,
    despesasMensais,
    cartoes,
    assinaturas,
    despesasTotais,
    parcelasFuturas,
    sobra: receitas - despesasTotais,
    criadoEm: new Date().toLocaleString("pt-BR"),
  };
}

function aplicarViradaDoMesNosDados(dadosOriginais: AppData): AppData {
  const dados = normalizarDadosOnline(dadosOriginais);
  const mesAtual = chaveMesAtual();
  const ultimaAtualizacao = dados.ultimaAtualizacaoMensal || mesAtual;
  const mesesPassados = diferencaMeses(ultimaAtualizacao, mesAtual);

  if (mesesPassados <= 0) {
    return dados;
  }

  const historicoDoMesAnterior = calcularHistoricoPorDados(
    ultimaAtualizacao,
    dados,
  );

  const comprasAtualizadas = dados.compras
    .map((item) => {
      if (tipoDaCompra(item) !== "parcela") return item;

      return {
        ...item,
        parcelaAtual: (item.parcelaAtual || 1) + mesesPassados,
      };
    })
    .filter((item) => {
      if (tipoDaCompra(item) !== "parcela") return true;
      return (item.parcelaAtual || 1) <= item.parcelas;
    });

  return {
    ...dados,
    compras: comprasAtualizadas,
    historicoMensal: [
      historicoDoMesAnterior,
      ...dados.historicoMensal.filter((item) => item.mes !== ultimaAtualizacao),
    ],
    ultimaAtualizacaoMensal: mesAtual,
  };
}

function formatar(valor: number) {
  return valor.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });
}

function numero(valor: string) {
  return Number(valor.replace(",", ".")) || 0;
}


function chaveMesAtual() {
  const hoje = new Date();
  const ano = hoje.getFullYear();
  const mes = String(hoje.getMonth() + 1).padStart(2, "0");
  return `${ano}-${mes}`;
}

function diferencaMeses(mesInicial: string, mesFinal: string) {
  const [anoInicial, mesInicialNumero] = mesInicial.split("-").map(Number);
  const [anoFinal, mesFinalNumero] = mesFinal.split("-").map(Number);

  if (!anoInicial || !mesInicialNumero || !anoFinal || !mesFinalNumero) {
    return 0;
  }

  return (anoFinal - anoInicial) * 12 + (mesFinalNumero - mesInicialNumero);
}

function tipoDaCompra(compra: Compra) {
  return compra.tipo || (compra.parcelas > 1 ? "parcela" : "avulso");
}

function nomeMesHistorico(mes: string) {
  const [ano, mesNumero] = mes.split("-").map(Number);

  if (!ano || !mesNumero) {
    return mes;
  }

  return new Date(ano, mesNumero - 1, 1).toLocaleDateString("pt-BR", {
    month: "long",
    year: "numeric",
  });
}

function limitarDiaDoMes(ano: number, mes: number, dia: number) {
  const ultimoDia = new Date(ano, mes + 1, 0).getDate();
  return Math.min(Math.max(Number(dia) || 1, 1), ultimoDia);
}

function proximaDataPorDia(dia: number) {
  const hoje = new Date();
  const ano = hoje.getFullYear();
  const mes = hoje.getMonth();
  const diaAjustado = limitarDiaDoMes(ano, mes, dia);
  let data = new Date(ano, mes, diaAjustado);

  if (data < new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate())) {
    const proximoMes = mes + 1;
    const anoProximo = new Date(ano, proximoMes, 1).getFullYear();
    const mesProximo = new Date(ano, proximoMes, 1).getMonth();
    data = new Date(anoProximo, mesProximo, limitarDiaDoMes(anoProximo, mesProximo, dia));
  }

  return data;
}

function diasAte(data: Date) {
  const hoje = new Date();
  const hojeLimpo = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  const dataLimpa = new Date(data.getFullYear(), data.getMonth(), data.getDate());
  return Math.ceil((dataLimpa.getTime() - hojeLimpo.getTime()) / (1000 * 60 * 60 * 24));
}

function formatarDataCurta(data: Date) {
  return data.toLocaleDateString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
  });
}

function lerJsonLocal<T>(chave: string, padrao: T): T {
  try {
    const valor = localStorage.getItem(chave);
    return valor ? JSON.parse(valor) : padrao;
  } catch {
    return padrao;
  }
}

function calcularHistoricoMensal(mes: string, comprasBase: Compra[]): HistoricoMensal {
  const receitasCltSalvas = lerJsonLocal<Lancamento[]>(
    "receitasClt",
    receitasCltIniciais,
  );
  const receitasMeiSalvas = lerJsonLocal<Lancamento[]>(
    "receitasMei",
    receitasMeiIniciais,
  );
  const despesasMensaisSalvas = lerJsonLocal<Lancamento[]>(
    "despesasMensais",
    despesasMensaisIniciais,
  );
  const assinaturasSalvas = lerJsonLocal<Assinatura[]>("assinaturas", []);

  const receitas =
    receitasCltSalvas.reduce((soma, item) => soma + item.valor, 0) +
    receitasMeiSalvas.reduce((soma, item) => soma + item.valor, 0);

  const despesasMensais = despesasMensaisSalvas.reduce(
    (soma, item) => soma + item.valor,
    0,
  );

  const cartoes = comprasBase
    .filter((item) => item.pessoa === "Rennan")
    .reduce((soma, item) => soma + item.valor, 0);

  const assinaturas = assinaturasSalvas
    .filter((item) => (item.pessoa || "Rennan") === "Rennan")
    .reduce((soma, item) => soma + item.valor, 0);

  const despesasTotais = despesasMensais + cartoes + assinaturas;

  const parcelasFuturas = comprasBase
    .filter((item) => item.pessoa === "Rennan")
    .reduce((soma, item) => {
      if (tipoDaCompra(item) !== "parcela") {
        return soma;
      }

      const atual = item.parcelaAtual || 1;
      const restantesIncluindoAtual = Math.max(item.parcelas - atual + 1, 0);

      return soma + item.valor * restantesIncluindoAtual;
    }, 0);

  return {
    mes,
    receitas,
    despesasMensais,
    cartoes,
    assinaturas,
    despesasTotais,
    parcelasFuturas,
    sobra: receitas - despesasTotais,
    criadoEm: new Date().toLocaleString("pt-BR"),
  };
}

function salvarHistoricoDoMes(mes: string, comprasBase: Compra[]) {
  const historicoAtual = lerJsonLocal<HistoricoMensal[]>("historicoMensal", []);
  const novoRegistro = calcularHistoricoMensal(mes, comprasBase);
  const historicoAtualizado = [
    novoRegistro,
    ...historicoAtual.filter((item) => item.mes !== mes),
  ];

  localStorage.setItem("historicoMensal", JSON.stringify(historicoAtualizado));
}

function atualizarParcelasNaViradaDoMes(compras: Compra[]) {
  const mesAtual = chaveMesAtual();
  const ultimaAtualizacao = localStorage.getItem("ultimaAtualizacaoMensal");

  if (!ultimaAtualizacao) {
    localStorage.setItem("ultimaAtualizacaoMensal", mesAtual);
    return compras;
  }

  const mesesPassados = diferencaMeses(ultimaAtualizacao, mesAtual);

  if (mesesPassados <= 0) {
    return compras;
  }

  salvarHistoricoDoMes(ultimaAtualizacao, compras);

  const comprasAtualizadas = compras
    .map((item) => {
      if (tipoDaCompra(item) !== "parcela") {
        return item;
      }

      const parcelaAtualNova = (item.parcelaAtual || 1) + mesesPassados;

      return {
        ...item,
        parcelaAtual: parcelaAtualNova,
      };
    })
    .filter((item) => {
      if (tipoDaCompra(item) !== "parcela") {
        return true;
      }

      return (item.parcelaAtual || 1) <= item.parcelas;
    });

  localStorage.setItem("ultimaAtualizacaoMensal", mesAtual);
  localStorage.setItem("compras", JSON.stringify(comprasAtualizadas));

  return comprasAtualizadas;
}

export default function Home() {
  const [aba, setAba] = useState("dashboard");
  const [subAbaMensal, setSubAbaMensal] = useState<"receitas" | "despesas">(
    "receitas",
  );

  const [compras, setCompras] = useState<Compra[]>([]);
  const [cartoes, setCartoes] = useState<Cartao[]>(cartoesIniciais);
  const [faturasCartoes, setFaturasCartoes] = useState<FaturaCartao[]>(faturasCartoesIniciais);
  const [pessoas, setPessoas] = useState<string[]>(pessoasIniciais);
  const [pessoasSelecionadasCartoes, setPessoasSelecionadasCartoes] = useState<
    string[]
  >(["Rennan"]);
  const [pessoasSelecionadasParcelas, setPessoasSelecionadasParcelas] =
    useState<string[]>(["Rennan"]);
  const [pessoaDetalhes, setPessoaDetalhes] = useState<string | null>(null);
  const [cartaoAvulsoAberto, setCartaoAvulsoAberto] = useState<string | null>(null);
  const [cartaoAssinaturasAberto, setCartaoAssinaturasAberto] = useState<string | null>(null);
  const [faturaDetalhesAberta, setFaturaDetalhesAberta] = useState<string | null>(null);
  const [cartaoVencimentoAberto, setCartaoVencimentoAberto] = useState<string | null>(null);
  const [assinaturas, setAssinaturas] = useState<Assinatura[]>([]);
  const [assinaturaEditandoId, setAssinaturaEditandoId] = useState<number | null>(null);
  const [assinaturaDescricao, setAssinaturaDescricao] = useState("");
  const [assinaturaValor, setAssinaturaValor] = useState("");
  const [assinaturaCartao, setAssinaturaCartao] = useState("Platinum");
  const [assinaturaPessoa, setAssinaturaPessoa] = useState("Rennan");

  const [receitasClt, setReceitasClt] =
    useState<Lancamento[]>(receitasCltIniciais);
  const [receitasMei, setReceitasMei] =
    useState<Lancamento[]>(receitasMeiIniciais);
  const [despesasMensais, setDespesasMensais] = useState<Lancamento[]>(
    despesasMensaisIniciais,
  );
  const [historicoMensal, setHistoricoMensal] = useState<HistoricoMensal[]>([]);
  const [ultimaAtualizacaoMensal, setUltimaAtualizacaoMensal] =
    useState(chaveMesAtual());
  const [dadosCarregados, setDadosCarregados] = useState(false);
  const [carregandoDados, setCarregandoDados] = useState(true);

  const [modalAberto, setModalAberto] = useState(false);
  const [compraEditandoId, setCompraEditandoId] = useState<number | null>(null);
  const [descricao, setDescricao] = useState("");
  const [valor, setValor] = useState("");
  const [cartao, setCartao] = useState("Platinum");
  const [pessoa, setPessoa] = useState("Rennan");
  const [parcelas, setParcelas] = useState("1");
  const [parcelaAtual, setParcelaAtual] = useState("1");

  const [novaPessoa, setNovaPessoa] = useState("");
  const [novoCartaoNome, setNovoCartaoNome] = useState("");
  const [novoCartaoLimite, setNovoCartaoLimite] = useState("");

  const [novaCltNome, setNovaCltNome] = useState("");
  const [novaCltValor, setNovaCltValor] = useState("");
  const [novaMeiNome, setNovaMeiNome] = useState("");
  const [novaMeiValor, setNovaMeiValor] = useState("");

  const [novaDespesaNome, setNovaDespesaNome] = useState("");
  const [novaDespesaValor, setNovaDespesaValor] = useState("");

  useEffect(() => {
    if (typeof navigator !== "undefined" && "serviceWorker" in navigator) {
      navigator.serviceWorker.register("/sw.js").catch(() => {});
    }

    async function carregarDadosOnline() {
      setCarregandoDados(true);

      const { data: registro, error } = await supabase
        .from("app_data")
        .select("data")
        .eq("id", APP_DATA_ID)
        .maybeSingle();

      if (error) {
        console.error("Erro ao carregar dados do Supabase:", error);
        alert("Erro ao carregar dados online. Verifique o Supabase.");
      }

      const dadosOnline = registro?.data as Partial<AppData> | null;
      const dadosBase =
        dadosOnline && Object.keys(dadosOnline).length > 0
          ? normalizarDadosOnline(dadosOnline)
          : montarDadosDoLocalStorage();

      const dadosAtualizados = aplicarViradaDoMesNosDados(dadosBase);

      setCompras(dadosAtualizados.compras);
      setCartoes(dadosAtualizados.cartoes);
      setFaturasCartoes(dadosAtualizados.faturasCartoes);
      setPessoas(dadosAtualizados.pessoas);
      setPessoasSelecionadasCartoes(dadosAtualizados.pessoasSelecionadasCartoes);
      setPessoasSelecionadasParcelas(dadosAtualizados.pessoasSelecionadasParcelas);
      setReceitasClt(dadosAtualizados.receitasClt);
      setReceitasMei(dadosAtualizados.receitasMei);
      setDespesasMensais(dadosAtualizados.despesasMensais);
      setAssinaturas(dadosAtualizados.assinaturas);
      setHistoricoMensal(dadosAtualizados.historicoMensal);
      setUltimaAtualizacaoMensal(dadosAtualizados.ultimaAtualizacaoMensal);

      await supabase.from("app_data").upsert({
        id: APP_DATA_ID,
        data: dadosAtualizados,
        updated_at: new Date().toISOString(),
      });

      setDadosCarregados(true);
      setCarregandoDados(false);
    }

    carregarDadosOnline();
  }, []);

  useEffect(() => {
    if (!dadosCarregados) return;

    const dadosAtualizados: AppData = {
      compras,
      cartoes,
      faturasCartoes,
      pessoas,
      pessoasSelecionadasCartoes,
      pessoasSelecionadasParcelas,
      receitasClt,
      receitasMei,
      despesasMensais,
      assinaturas,
      historicoMensal,
      ultimaAtualizacaoMensal,
    };

    const timeout = window.setTimeout(async () => {
      const { error } = await supabase.from("app_data").upsert({
        id: APP_DATA_ID,
        data: dadosAtualizados,
        updated_at: new Date().toISOString(),
      });

      if (error) {
        console.error("Erro ao salvar dados no Supabase:", error);
      }
    }, 500);

    return () => window.clearTimeout(timeout);
  }, [
    dadosCarregados,
    compras,
    cartoes,
    faturasCartoes,
    pessoas,
    pessoasSelecionadasCartoes,
    pessoasSelecionadasParcelas,
    receitasClt,
    receitasMei,
    despesasMensais,
    assinaturas,
    historicoMensal,
    ultimaAtualizacaoMensal,
  ]);

  const totalReceitasClt = useMemo(
    () => receitasClt.reduce((soma, item) => soma + item.valor, 0),
    [receitasClt],
  );

  const totalReceitasMei = useMemo(
    () => receitasMei.reduce((soma, item) => soma + item.valor, 0),
    [receitasMei],
  );

  const totalReceitas = totalReceitasClt + totalReceitasMei;

  const totalDespesasMensais = useMemo(
    () => despesasMensais.reduce((soma, item) => soma + item.valor, 0),
    [despesasMensais],
  );

  const totalAssinaturas = useMemo(
    () =>
      assinaturas
        .filter((item) => item.pessoa === "Rennan")
        .reduce((soma, item) => soma + item.valor, 0),
    [assinaturas],
  );

  const totalComprasCartao = useMemo(
    () =>
      compras
        .filter((item) => item.pessoa === "Rennan")
        .reduce((soma, item) => soma + item.valor, 0),
    [compras],
  );

  const totalDespesas = totalDespesasMensais + totalComprasCartao + totalAssinaturas;
  const sobra = totalReceitas - totalDespesas;

  const totalParcelasFuturas = useMemo(
    () =>
      compras
        .filter((item) => item.pessoa === "Rennan")
        .reduce((soma, item) => {
          if ((item.tipo || (item.parcelas > 1 ? "parcela" : "avulso")) !== "parcela") {
            return soma;
          }

          const parcelaAtualCompra = item.parcelaAtual || 1;
          const restantesIncluindoAtual = Math.max(
            item.parcelas - parcelaAtualCompra + 1,
            0,
          );

          return soma + item.valor * restantesIncluindoAtual;
        }, 0),
    [compras],
  );

  const comprasFiltradasCartoes = useMemo(() => {
    return compras.filter((item) => item.pessoa === "Rennan");
  }, [compras]);

  const totalCartoesFiltrado = useMemo(() => {
    const totalCompras = comprasFiltradasCartoes.reduce((soma, item) => {
      const tipo = item.tipo || (item.parcelas > 1 ? "parcela" : "avulso");

      if (tipo === "parcela") {
        const atual = item.parcelaAtual || 1;
        const restantesIncluindoAtual = Math.max(item.parcelas - atual + 1, 0);
        return soma + item.valor * restantesIncluindoAtual;
      }

      return soma + item.valor;
    }, 0);

    return totalCompras + totalAssinaturas;
  }, [comprasFiltradasCartoes, totalAssinaturas]);

  const gastosPorCartao = cartoes.map((cartaoItem) => {
    const gastoCompras = comprasFiltradasCartoes
      .filter((item) => item.cartao === cartaoItem.nome)
      .reduce((soma, item) => {
        const tipo = item.tipo || (item.parcelas > 1 ? "parcela" : "avulso");

        if (tipo === "parcela") {
          const atual = item.parcelaAtual || 1;
          const restantesIncluindoAtual = Math.max(item.parcelas - atual + 1, 0);
          return soma + item.valor * restantesIncluindoAtual;
        }

        return soma + item.valor;
      }, 0);

    const gastoAssinaturas = assinaturas
      .filter((item) => item.cartao === cartaoItem.nome && item.pessoa === "Rennan")
      .reduce((soma, item) => soma + item.valor, 0);

    const gasto = gastoCompras + gastoAssinaturas;

    const uso =
      cartaoItem.limite > 0
        ? Math.min((gasto / cartaoItem.limite) * 100, 100)
        : 0;

    return { nome: cartaoItem.nome, gasto, limite: cartaoItem.limite, uso };
  });

  const faturasResumo = faturasCartoes.map((fatura) => {
    const dataVencimento = proximaDataPorDia(fatura.vencimento);
    const dias = diasAte(dataVencimento);
    const totalCartao = gastosPorCartao.find((item) => item.nome === fatura.nome)?.gasto || 0;

    return {
      ...fatura,
      dataVencimento,
      dias,
      totalCartao,
    };
  });

  const proximaFatura = [...faturasResumo].sort((a, b) => a.dias - b.dias)[0];

  function alterarFaturaCartao(nome: string, campo: "fechamento" | "vencimento", valorNovo: string) {
    const valorDia = Math.min(Math.max(Number(valorNovo) || 1, 1), 31);

    setFaturasCartoes((listaAtual) => {
      const existe = listaAtual.some((item) => item.nome === nome);
      const listaBase = existe
        ? listaAtual
        : [...listaAtual, { nome, fechamento: 1, vencimento: 1 }];

      return listaBase.map((item) =>
        item.nome === nome ? { ...item, [campo]: valorDia } : item,
      );
    });
  }

  const comprasFiltradasParcelas = useMemo(() => {
    if (pessoasSelecionadasParcelas.length === 0) return [];
    return compras.filter((item) =>
      pessoasSelecionadasParcelas.includes(item.pessoa),
    );
  }, [compras, pessoasSelecionadasParcelas]);

  const parcelasPorCartao = cartoes.map((cartaoItem) => {
    const itens = comprasFiltradasParcelas.filter(
      (item) =>
        item.cartao === cartaoItem.nome &&
        item.parcelas > 1 &&
        (item.tipo || (item.parcelas > 1 ? "parcela" : "avulso")) === "parcela",
    );

    const totalParcelas = itens.reduce((soma, item) => {
      const atual = item.parcelaAtual || 1;
      const restantesIncluindoAtual = Math.max(item.parcelas - atual + 1, 0);
      return soma + item.valor * restantesIncluindoAtual;
    }, 0);

    return {
      nome: cartaoItem.nome,
      itens,
      totalParcelas,
    };
  });

  const totalParcelasSelecionadas = parcelasPorCartao.reduce(
    (soma, cartao) => soma + cartao.totalParcelas,
    0,
  );

  function limparFormularioCompra() {
    setDescricao("");
    setValor("");
    setCartao(cartoes[0]?.nome || "");
    setPessoa(pessoas[0] || "");
    setParcelas("1");
    setParcelaAtual("1");
    setCompraEditandoId(null);
  }

  function abrirModalNovaCompra() {
    limparFormularioCompra();
    setModalAberto(true);
  }

  function abrirModalEditarCompra(compra: Compra) {
    setCompraEditandoId(compra.id);
    setDescricao(compra.descricao);
    setValor(String(compra.valor));
    setCartao(compra.cartao);
    setPessoa(compra.pessoa);
    setParcelas(String(compra.parcelas));
    setParcelaAtual(String(compra.parcelaAtual || 1));
    setModalAberto(true);
  }

  function salvarCompra() {
    if (!descricao || !valor) return;

    if (compraEditandoId) {
      setCompras(
        compras.map((item) =>
          item.id === compraEditandoId
            ? {
                ...item,
                descricao,
                valor: numero(valor),
                cartao,
                pessoa,
                parcelas: Number(parcelas) || 1,
                parcelaAtual: Number(parcelaAtual) || 1,
                tipo: (Number(parcelas) || 1) > 1 ? "parcela" : "avulso",
              }
            : item,
        ),
      );
    } else {
      const novaCompra: Compra = {
        id: Date.now(),
        descricao,
        valor: numero(valor),
        cartao,
        pessoa,
        parcelas: Number(parcelas) || 1,
        parcelaAtual: Number(parcelaAtual) || 1,
        tipo: (Number(parcelas) || 1) > 1 ? "parcela" : "avulso",
        data: new Date().toLocaleDateString("pt-BR"),
      };

      setCompras([novaCompra, ...compras]);
    }

    limparFormularioCompra();
    setModalAberto(false);
  }

  function salvarGastoAvulso() {
    if (!descricao || !valor) return;

    const novoGasto: Compra = {
      id: Date.now(),
      descricao,
      valor: numero(valor),
      cartao,
      pessoa,
      parcelas: 1,
      parcelaAtual: 1,
      tipo: "avulso",
      data: new Date().toLocaleDateString("pt-BR"),
    };

    setCompras([novoGasto, ...compras]);
    limparFormularioCompra();
  }

  function limparFormularioAssinatura(cartaoAtual?: string) {
    setAssinaturaDescricao("");
    setAssinaturaValor("");
    setAssinaturaCartao(cartaoAtual || cartoes[0]?.nome || "");
    setAssinaturaPessoa("Rennan");
    setAssinaturaEditandoId(null);
  }

  function salvarAssinatura(cartaoAtual: string) {
    if (!assinaturaDescricao || !assinaturaValor) return;

    if (assinaturaEditandoId) {
      setAssinaturas(
        assinaturas.map((item) =>
          item.id === assinaturaEditandoId
            ? {
                ...item,
                descricao: assinaturaDescricao,
                valor: numero(assinaturaValor),
                cartao: assinaturaCartao || cartaoAtual,
                pessoa: assinaturaPessoa || "Rennan",
              }
            : item,
        ),
      );
    } else {
      const novaAssinatura: Assinatura = {
        id: Date.now(),
        descricao: assinaturaDescricao,
        valor: numero(assinaturaValor),
        cartao: cartaoAtual,
        pessoa: assinaturaPessoa || "Rennan",
      };

      setAssinaturas([novaAssinatura, ...assinaturas]);
    }

    limparFormularioAssinatura(cartaoAtual);
  }

  function abrirEditarAssinatura(assinatura: Assinatura) {
    setAssinaturaEditandoId(assinatura.id);
    setAssinaturaDescricao(assinatura.descricao);
    setAssinaturaValor(String(assinatura.valor));
    setAssinaturaCartao(assinatura.cartao);
    setAssinaturaPessoa(assinatura.pessoa || "Rennan");
  }

  function excluirAssinatura(id: number) {
    setAssinaturas(assinaturas.filter((item) => item.id !== id));
  }

  function excluirCompra(id: number) {
    setCompras(compras.filter((item) => item.id !== id));
  }

  function editarCompra(
    id: number,
    campo:
      | "descricao"
      | "valor"
      | "cartao"
      | "pessoa"
      | "parcelaAtual"
      | "parcelas",
    valorNovo: string,
  ) {
    setCompras(
      compras.map((item) => {
        if (item.id !== id) return item;

        if (["valor", "parcelaAtual", "parcelas"].includes(campo)) {
          return { ...item, [campo]: numero(valorNovo) };
        }

        return { ...item, [campo]: valorNovo };
      }),
    );
  }

  function alterarLimite(nome: string, novoLimite: string) {
    setCartoes(
      cartoes.map((cartao) =>
        cartao.nome === nome
          ? { ...cartao, limite: numero(novoLimite) }
          : cartao,
      ),
    );
  }

  function adicionarCartao() {
    const nome = novoCartaoNome.trim();

    if (!nome) return;

    const existe = cartoes.some(
      (item) => item.nome.toLowerCase() === nome.toLowerCase(),
    );

    if (existe) {
      alert("Esse cartão já existe.");
      return;
    }

    const limite = numero(novoCartaoLimite) || 0;

    setCartoes([...cartoes, { nome, limite }]);
    setFaturasCartoes([
      ...faturasCartoes,
      { nome, fechamento: 1, vencimento: 1 },
    ]);

    if (!cartao) {
      setCartao(nome);
    }

    setNovoCartaoNome("");
    setNovoCartaoLimite("");
  }

  function excluirCartao(nome: string) {
    const comprasDoCartao = compras.filter((item) => item.cartao === nome);
    const assinaturasDoCartao = assinaturas.filter((item) => item.cartao === nome);

    const mensagem =
      comprasDoCartao.length > 0 || assinaturasDoCartao.length > 0
        ? `Excluir o cartão ${nome} também vai remover ${comprasDoCartao.length} compra(s) e ${assinaturasDoCartao.length} assinatura(s) dele. Deseja continuar?`
        : `Deseja excluir o cartão ${nome}?`;

    const confirmar = window.confirm(mensagem);

    if (!confirmar) return;

    const novosCartoes = cartoes.filter((item) => item.nome !== nome);

    setCartoes(novosCartoes);
    setFaturasCartoes(faturasCartoes.filter((item) => item.nome !== nome));
    setCompras(compras.filter((item) => item.cartao !== nome));
    setAssinaturas(assinaturas.filter((item) => item.cartao !== nome));

    if (cartao === nome) {
      setCartao(novosCartoes[0]?.nome || "");
    }

    if (assinaturaCartao === nome) {
      setAssinaturaCartao(novosCartoes[0]?.nome || "");
    }

    if (cartaoAvulsoAberto === nome) setCartaoAvulsoAberto(null);
    if (cartaoAssinaturasAberto === nome) setCartaoAssinaturasAberto(null);
    if (faturaDetalhesAberta === nome) setFaturaDetalhesAberta(null);
    if (cartaoVencimentoAberto === nome) setCartaoVencimentoAberto(null);
  }

  function adicionarPessoa() {
    const nome = novaPessoa.trim();
    if (!nome) return;

    const existe = pessoas.some(
      (item) => item.toLowerCase() === nome.toLowerCase(),
    );

    if (existe) {
      alert("Essa pessoa já existe.");
      return;
    }

    setPessoas([...pessoas, nome]);
    setNovaPessoa("");
  }

  function excluirPessoa(nome: string) {
    const pessoaTemCompra = compras.some((item) => item.pessoa === nome);

    if (pessoaTemCompra) {
      alert(
        "Não é possível excluir essa pessoa porque ela possui compras cadastradas.",
      );
      return;
    }

    setPessoas(pessoas.filter((item) => item !== nome));
    setPessoasSelecionadasCartoes(
      pessoasSelecionadasCartoes.filter((item) => item !== nome),
    );
    setPessoasSelecionadasParcelas(
      pessoasSelecionadasParcelas.filter((item) => item !== nome),
    );
  }

  function alternarPessoaCartoes(nome: string) {
    setPessoasSelecionadasCartoes((listaAtual) => {
      if (listaAtual.includes(nome)) {
        return listaAtual.filter((item) => item !== nome);
      }

      return [...listaAtual, nome];
    });
  }

  function selecionarTodasPessoasCartoes() {
    if (pessoasSelecionadasCartoes.length === pessoas.length) {
      setPessoasSelecionadasCartoes([]);
      return;
    }

    setPessoasSelecionadasCartoes(pessoas);
  }

  function alternarPessoaParcelas(nome: string) {
    setPessoasSelecionadasParcelas((listaAtual) => {
      if (listaAtual.includes(nome)) {
        return listaAtual.filter((item) => item !== nome);
      }

      return [...listaAtual, nome];
    });
  }

  function selecionarTodasPessoasParcelas() {
    if (pessoasSelecionadasParcelas.length === pessoas.length) {
      setPessoasSelecionadasParcelas([]);
      return;
    }

    setPessoasSelecionadasParcelas(pessoas);
  }

  function comprasDaPessoa(nome: string) {
    return compras.filter((item) => item.pessoa === nome);
  }

  function assinaturasDaPessoa(nome: string) {
    return assinaturas.filter((item) => item.pessoa === nome);
  }

  function totalAssinaturasDaPessoa(nome: string) {
    return assinaturasDaPessoa(nome).reduce((soma, item) => soma + item.valor, 0);
  }

  function totalDaPessoa(nome: string) {
    const totalCompras = comprasDaPessoa(nome).reduce((soma, item) => soma + item.valor, 0);
    return totalCompras + totalAssinaturasDaPessoa(nome);
  }

  function totalParcelasDaPessoa(nome: string) {
    return comprasDaPessoa(nome).reduce((soma, item) => {
      const atual = item.parcelaAtual || 1;
      const restantesIncluindoAtual = Math.max(item.parcelas - atual + 1, 0);
      return soma + item.valor * restantesIncluindoAtual;
    }, 0);
  }


  function linhasDetalhesPessoa(nome: string) {
    const linhasCompras = comprasDaPessoa(nome).map((item) => {
      const atual = item.parcelaAtual || 1;
      const restantes = Math.max(item.parcelas - atual + 1, 0);
      const totalParcelas = item.valor * restantes;

      return {
        Compra: item.descricao,
        Tipo: item.tipo || (item.parcelas > 1 ? "parcela" : "avulso"),
        Valor_da_parcela: item.valor,
        Parcela: `${atual}/${item.parcelas}`,
        Parcelas_restantes: restantes,
        Total_de_parcelas: totalParcelas,
        Data: item.data,
      };
    });

    const linhasAssinaturas = assinaturasDaPessoa(nome).map((item) => ({
      Compra: item.descricao,
      Tipo: "assinatura",
      Valor_da_parcela: item.valor,
      Parcela: "fixa",
      Parcelas_restantes: 1,
      Total_de_parcelas: item.valor,
      Data: "Todo mês",
    }));

    return [...linhasCompras, ...linhasAssinaturas];
  }

  function baixarExcelPessoa(nome: string) {
    const linhas = linhasDetalhesPessoa(nome);

    const resumo = [
      ["Pessoa", nome],
      ["Total atual", totalDaPessoa(nome)],
      ["Total de parcelas", totalParcelasDaPessoa(nome)],
      ["Total de assinaturas", totalAssinaturasDaPessoa(nome)],
      [],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(resumo);
    XLSX.utils.sheet_add_json(worksheet, linhas, {
      origin: "A5",
      skipHeader: false,
    });

    worksheet["!cols"] = [
      { wch: 35 },
      { wch: 12 },
      { wch: 18 },
      { wch: 14 },
      { wch: 18 },
      { wch: 20 },
      { wch: 14 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Detalhes");
    XLSX.writeFile(workbook, `detalhes-${nome}.xlsx`);
  }

  function baixarPdfPessoa(nome: string) {
    const doc = new jsPDF();
    const linhas = linhasDetalhesPessoa(nome);

    doc.setFontSize(18);
    doc.text(`Detalhes de ${nome}`, 14, 18);

    doc.setFontSize(11);
    doc.text(`Total atual: ${formatar(totalDaPessoa(nome))}`, 14, 28);
    doc.text(`Total de parcelas: ${formatar(totalParcelasDaPessoa(nome))}`, 14, 35);
    doc.text(`Total de assinaturas: ${formatar(totalAssinaturasDaPessoa(nome))}`, 14, 42);

    let y = 55;

    doc.setFontSize(9);
    doc.text("Compra", 14, y);
    doc.text("Parcela", 80, y);
    doc.text("Rest.", 105, y);
    doc.text("Valor", 125, y);
    doc.text("Total", 155, y);

    y += 6;
    doc.line(14, y - 3, 196, y - 3);

    linhas.forEach((item) => {
      if (y > 280) {
        doc.addPage();
        y = 18;
      }

      const compra = String(item.Compra).length > 24
        ? `${String(item.Compra).slice(0, 24)}...`
        : String(item.Compra);

      doc.text(compra, 14, y);
      doc.text(String(item.Parcela), 80, y);
      doc.text(String(item.Parcelas_restantes), 105, y);
      doc.text(formatar(Number(item.Valor_da_parcela)), 125, y);
      doc.text(formatar(Number(item.Total_de_parcelas)), 155, y);

      y += 7;
    });

    y += 4;
    if (y > 280) {
      doc.addPage();
      y = 18;
    }

    doc.line(14, y - 3, 196, y - 3);
    doc.setFontSize(11);
    doc.text(`Total de parcelas: ${formatar(totalParcelasDaPessoa(nome))}`, 14, y + 5);

    doc.save(`detalhes-${nome}.pdf`);
  }

  function salvarHistoricoMesAtual() {
    const mesAtual = chaveMesAtual();

    const novoRegistro: HistoricoMensal = {
      mes: mesAtual,
      receitas: totalReceitas,
      despesasMensais: totalDespesasMensais,
      cartoes: totalComprasCartao,
      assinaturas: totalAssinaturas,
      despesasTotais: totalDespesas,
      parcelasFuturas: totalParcelasFuturas,
      sobra,
      criadoEm: new Date().toLocaleString("pt-BR"),
    };

    setHistoricoMensal((listaAtual) => [
      novoRegistro,
      ...listaAtual.filter((item) => item.mes !== mesAtual),
    ]);

    alert("Histórico do mês salvo com sucesso!");
  }

  function excluirHistorico(mes: string) {
    setHistoricoMensal(historicoMensal.filter((item) => item.mes !== mes));
  }


  function baixarBackup() {
    const backup = {
      versao: 1,
      criadoEm: new Date().toLocaleString("pt-BR"),
      ultimaAtualizacaoMensal,
      dados: {
        compras,
        cartoes,
        faturasCartoes,
        pessoas,
        pessoasSelecionadasCartoes,
        pessoasSelecionadasParcelas,
        receitasClt,
        receitasMei,
        despesasMensais,
        assinaturas,
        historicoMensal,
      },
    };

    const blob = new Blob([JSON.stringify(backup, null, 2)], {
      type: "application/json",
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `backup-controle-financeiro-${chaveMesAtual()}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  async function restaurarBackup(arquivo?: File) {
    if (!arquivo) return;

    const confirmar = window.confirm(
      "Restaurar este backup vai substituir todos os dados atuais do sistema. Deseja continuar?",
    );

    if (!confirmar) return;

    try {
      const textoArquivo = await arquivo.text();
      const backup = JSON.parse(textoArquivo);
      const dados = backup?.dados;

      if (!dados) {
        alert("Arquivo de backup inválido.");
        return;
      }

      const comprasRestauradas: Compra[] = (dados.compras || []).map((item: Compra) => ({
        ...item,
        tipo: item.tipo || (item.parcelas > 1 ? "parcela" : "avulso"),
      }));

      const assinaturasRestauradas: Assinatura[] = (dados.assinaturas || []).map(
        (item: Assinatura) => ({
          ...item,
          pessoa: item.pessoa || "Rennan",
        }),
      );

      setCompras(comprasRestauradas);
      setCartoes(dados.cartoes || cartoesIniciais);
      setFaturasCartoes(dados.faturasCartoes || faturasCartoesIniciais);
      setPessoas(dados.pessoas || pessoasIniciais);
      setPessoasSelecionadasCartoes(dados.pessoasSelecionadasCartoes || ["Rennan"]);
      setPessoasSelecionadasParcelas(dados.pessoasSelecionadasParcelas || ["Rennan"]);
      setReceitasClt(dados.receitasClt || receitasCltIniciais);
      setReceitasMei(dados.receitasMei || receitasMeiIniciais);
      setDespesasMensais(dados.despesasMensais || despesasMensaisIniciais);
      setAssinaturas(assinaturasRestauradas);
      setHistoricoMensal(dados.historicoMensal || []);

      setUltimaAtualizacaoMensal(backup.ultimaAtualizacaoMensal || chaveMesAtual());

      alert("Backup restaurado com sucesso!");
    } catch {
      alert("Não foi possível restaurar o backup. Verifique se o arquivo está correto.");
    }
  }

  function adicionarReceitaClt() {
    if (!novaCltNome || !novaCltValor) return;

    setReceitasClt([
      ...receitasClt,
      { id: Date.now(), descricao: novaCltNome, valor: numero(novaCltValor) },
    ]);

    setNovaCltNome("");
    setNovaCltValor("");
  }

  function adicionarReceitaMei() {
    if (!novaMeiNome || !novaMeiValor) return;

    setReceitasMei([
      ...receitasMei,
      { id: Date.now(), descricao: novaMeiNome, valor: numero(novaMeiValor) },
    ]);

    setNovaMeiNome("");
    setNovaMeiValor("");
  }

  function adicionarDespesaMensal() {
    if (!novaDespesaNome || !novaDespesaValor) return;

    setDespesasMensais([
      ...despesasMensais,
      {
        id: Date.now(),
        descricao: novaDespesaNome,
        valor: numero(novaDespesaValor),
      },
    ]);

    setNovaDespesaNome("");
    setNovaDespesaValor("");
  }

  function editarReceitaClt(
    id: number,
    campo: "descricao" | "valor",
    valorNovo: string,
  ) {
    setReceitasClt(
      receitasClt.map((item) =>
        item.id === id
          ? {
              ...item,
              [campo]: campo === "valor" ? numero(valorNovo) : valorNovo,
            }
          : item,
      ),
    );
  }

  function editarReceitaMei(
    id: number,
    campo: "descricao" | "valor",
    valorNovo: string,
  ) {
    setReceitasMei(
      receitasMei.map((item) =>
        item.id === id
          ? {
              ...item,
              [campo]: campo === "valor" ? numero(valorNovo) : valorNovo,
            }
          : item,
      ),
    );
  }

  function editarDespesaMensal(
    id: number,
    campo: "descricao" | "valor",
    valorNovo: string,
  ) {
    setDespesasMensais(
      despesasMensais.map((item) =>
        item.id === id
          ? {
              ...item,
              [campo]: campo === "valor" ? numero(valorNovo) : valorNovo,
            }
          : item,
      ),
    );
  }

  async function importarPlanilha(arquivo?: File) {
    if (!arquivo) return;

    const confirmar = window.confirm(
      "Importar essa planilha vai substituir receitas, despesas mensais e compras atuais. Deseja continuar?",
    );

    if (!confirmar) return;

    const buffer = await arquivo.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });

    const linhas = (nomeAba: string) => {
      const sheet = workbook.Sheets[nomeAba];
      if (!sheet) return [] as unknown[][];
      return XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
      }) as unknown[][];
    };

    const texto = (valor: unknown) => String(valor ?? "").trim();
    const dinheiro = (valor: unknown) => {
      if (typeof valor === "number") return valor;
      return numero(
        String(valor ?? "")
          .replace("R$", "")
          .trim(),
      );
    };

    const mensal = linhas("Mensal");

    const cltImportado: Lancamento[] = [];
    const despesasImportadas: Lancamento[] = [];
    const meiImportado: Lancamento[] = [];

    mensal.forEach((linha, index) => {
      const nomeColA = texto(linha[0]);
      const valorColB = dinheiro(linha[1]);

      if (["Salário", "Salario", "Extra"].includes(nomeColA)) {
        cltImportado.push({
          id: Date.now() + index,
          descricao: nomeColA,
          valor: valorColB,
        });
      }

      if (
        nomeColA &&
        ![
          "Julho",
          "Salário",
          "Salario",
          "Extra",
          "Cartão",
          "Gastos",
          "Ganho",
          "Sobra",
        ].includes(nomeColA) &&
        typeof linha[1] !== "undefined" &&
        valorColB > 0
      ) {
        despesasImportadas.push({
          id: Date.now() + 1000 + index,
          descricao: nomeColA,
          valor: valorColB,
        });
      }

      const nomeMei = texto(linha[4]);
      const valorMei = dinheiro(linha[5]);

      if (nomeMei && !["MEI", "Total"].includes(nomeMei) && valorMei > 0) {
        meiImportado.push({
          id: Date.now() + 2000 + index,
          descricao: nomeMei,
          valor: valorMei,
        });
      }

      const nomePix = texto(linha[7]);
      const valorPix = dinheiro(linha[8]);

      if (
        nomePix &&
        ![
          "PIX/DEBITO",
          "PIX/DÉBITO",
          "Compra",
          "Total",
          "A vista mãe",
        ].includes(nomePix) &&
        valorPix > 0
      ) {
        despesasImportadas.push({
          id: Date.now() + 3000 + index,
          descricao: `PIX/Débito - ${nomePix}`,
          valor: valorPix,
        });
      }
    });

    const importarComprasCartao = (nomeAba: string, nomeCartao: string) => {
      const rows = linhas(nomeAba);
      if (!rows.length) return [] as Compra[];

      const cabecalho = rows[1] || [];
      const colunasPessoas: { indice: number; nome: string }[] = [];

      for (let coluna = 1; coluna < cabecalho.length; coluna++) {
        const nome = texto(cabecalho[coluna]);
        if (!nome || nome === "Total") break;
        colunasPessoas.push({ indice: coluna, nome });
      }

      const resultado: Compra[] = [];

      for (let linhaIndex = 2; linhaIndex < rows.length; linhaIndex++) {
        const linha = rows[linhaIndex];
        const descricaoCompra = texto(linha[0]);

        if (!descricaoCompra) continue;
        if (descricaoCompra.toLowerCase() === "total") break;

        const parcelasEncontradas = descricaoCompra.match(/\((\d+)\/(\d+)\)/);
        const parcelaAtualImportada = parcelasEncontradas
          ? Number(parcelasEncontradas[1])
          : 1;
        const totalParcelas = parcelasEncontradas
          ? Number(parcelasEncontradas[2])
          : 1;
        const descricaoLimpa = descricaoCompra
          .replace(/\s*\(\d+\/\d+\)\s*/g, "")
          .trim();

        colunasPessoas.forEach(({ indice, nome }) => {
          const valorCompra = dinheiro(linha[indice]);

          if (valorCompra > 0) {
            resultado.push({
              id:
                Date.now() +
                resultado.length +
                linhaIndex +
                Math.floor(Math.random() * 100000),
              descricao: descricaoLimpa || descricaoCompra,
              valor: valorCompra,
              cartao: nomeCartao,
              pessoa: nome,
              parcelas: totalParcelas || 1,
              parcelaAtual: parcelaAtualImportada || 1,
              tipo: (totalParcelas || 1) > 1 ? "parcela" : "avulso",
              data: new Date().toLocaleDateString("pt-BR"),
            });
          }
        });
      }

      return resultado;
    };

    const comprasImportadas = [
      ...importarComprasCartao("ITAU PLATINUM", "Platinum"),
      ...importarComprasCartao("ITAU UNICLASS", "Uniclass"),
      ...importarComprasCartao("LATAM", "Latam"),
      ...importarComprasCartao("MERCADO PAGO", "Mercado Pago"),
    ];

    const pessoasImportadas = Array.from(
      new Set(
        [...pessoas, ...comprasImportadas.map((item) => item.pessoa)].filter(
          Boolean,
        ),
      ),
    );

    setReceitasClt(cltImportado);
    setReceitasMei(meiImportado);
    setDespesasMensais(despesasImportadas);
    setCompras(comprasImportadas);
    setPessoas(pessoasImportadas);

    alert("Planilha importada com sucesso!");
  }

  if (carregandoDados) {
    return (
      <main className="min-h-screen bg-zinc-950 text-white flex items-center justify-center">
        <div className="rounded-xl border border-zinc-800 bg-zinc-900 p-8 text-center">
          <p className="text-2xl font-bold">Carregando dados online...</p>
          <p className="mt-2 text-zinc-400">Conectando ao Supabase</p>
        </div>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-zinc-950 text-white">
      <aside className="fixed left-0 top-0 h-full w-72 bg-zinc-900 border-r border-zinc-800 p-6">
        <h1 className="text-2xl font-bold">💰 Controle Financeiro</h1>
        <p className="text-zinc-500 text-sm mt-1 mb-8">Gestão pessoal</p>

        <nav className="space-y-3">
          {[
            ["dashboard", "📊 Dashboard"],
            ["mensal", "💰 Mensal"],
            ["cartoes", "💳 Cartões"],
            ["parcelas", "📅 Parcelas"],
            ["pessoas", "👥 Pessoas"],
            ["historico", "📈 Histórico"],
            ["config", "⚙️ Configurações"],
          ].map(([id, nome]) => (
            <button
              key={id}
              type="button"
              onClick={() => setAba(id)}
              className={`w-full text-left px-4 py-3 rounded-lg transition ${
                aba === id ? "bg-green-600" : "bg-zinc-800 hover:bg-zinc-700"
              }`}
            >
              {nome}
            </button>
          ))}
        </nav>
      </aside>

      <section className="ml-72 p-8">
        {aba === "dashboard" && (
          <>
            <h2 className="text-4xl font-bold mb-8">Dashboard</h2>

            <div className="grid grid-cols-4 gap-4 mb-8">
              <Card
                titulo="Receitas do mês"
                valor={formatar(totalReceitas)}
                cor="text-green-500"
              />
              <Card
                titulo="Despesas do mês"
                valor={formatar(totalDespesas)}
                cor="text-red-500"
              />
              <Card
                titulo="Sobra"
                valor={formatar(sobra)}
                cor={sobra >= 0 ? "text-blue-400" : "text-red-500"}
              />
              <Card
                titulo="Parcelas futuras"
                valor={formatar(totalParcelasFuturas)}
                cor="text-yellow-400"
              />
            </div>

            <div className="grid grid-cols-4 gap-4 mb-8">
              <Card
                titulo="CLT"
                valor={formatar(totalReceitasClt)}
                cor="text-green-400"
              />
              <Card
                titulo="MEI"
                valor={formatar(totalReceitasMei)}
                cor="text-emerald-400"
              />
              <Card
                titulo="Despesas mensais"
                valor={formatar(totalDespesasMensais)}
                cor="text-red-400"
              />
              <Card
                titulo="Assinaturas"
                valor={formatar(totalAssinaturas)}
                cor="text-orange-400"
              />
            </div>

            <div className="flex items-center justify-between mb-4">
              <div>
                <h3 className="text-2xl font-bold">Resumo por cartão</h3>
                <p className="text-zinc-400 text-sm mt-1">
                  Mostrando somente Rennan, com avulsos, assinaturas e todas as parcelas futuras.
                </p>
              </div>

              <div className="text-right">
                <p className="text-zinc-400 text-sm">Total dos cartões</p>
                <p className="text-2xl font-bold text-red-500">
                  {formatar(totalCartoesFiltrado)}
                </p>
              </div>
            </div>

            <div className="grid grid-cols-4 gap-4">
              {gastosPorCartao.map((cartao) => (
                <div
                  key={cartao.nome}
                  className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl"
                >
                  <p className="text-zinc-400">{cartao.nome}</p>
                  <p className="text-2xl font-bold">{formatar(cartao.gasto)}</p>
                  <p className="text-sm text-zinc-500 mt-2">
                    Limite: {formatar(cartao.limite)}
                  </p>
                  <p className="text-xs text-zinc-400 mt-1">
                    Fecha dia {faturasCartoes.find((item) => item.nome === cartao.nome)?.fechamento || "-"} • Vence dia {faturasCartoes.find((item) => item.nome === cartao.nome)?.vencimento || "-"}
                  </p>

                  <div className="w-full bg-zinc-800 rounded-full h-2 mt-4">
                    <div
                      className="bg-green-500 h-2 rounded-full"
                      style={{ width: `${cartao.uso}%` }}
                    />
                  </div>

                  <p className="text-sm text-zinc-400 mt-2">
                    Uso: {cartao.uso.toFixed(0)}%
                  </p>

                  <div className="mt-4 flex gap-2">
                    <button
                      type="button"
                      onClick={() => setFaturaDetalhesAberta(cartao.nome)}
                      className="w-full rounded-lg bg-zinc-800 px-4 py-2 text-sm font-bold text-zinc-200 hover:bg-zinc-700"
                    >
                      Ver fatura
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </>
        )}

        {aba === "mensal" && (
          <>
            <h2 className="text-4xl font-bold mb-8">Mensal</h2>

            <div className="flex gap-3 mb-6">
              <button
                type="button"
                onClick={() => setSubAbaMensal("receitas")}
                className={`px-5 py-3 rounded-lg font-bold ${
                  subAbaMensal === "receitas" ? "bg-green-600" : "bg-zinc-800"
                }`}
              >
                Receitas
              </button>

              <button
                type="button"
                onClick={() => setSubAbaMensal("despesas")}
                className={`px-5 py-3 rounded-lg font-bold ${
                  subAbaMensal === "despesas" ? "bg-red-600" : "bg-zinc-800"
                }`}
              >
                Despesas
              </button>
            </div>

            {subAbaMensal === "receitas" && (
              <>
                <div className="grid grid-cols-2 gap-6">
                  <TabelaLancamentos
                    titulo="CLT"
                    cor="green"
                    itens={receitasClt}
                    novoNome={novaCltNome}
                    novoValor={novaCltValor}
                    setNovoNome={setNovaCltNome}
                    setNovoValor={setNovaCltValor}
                    adicionar={adicionarReceitaClt}
                    editar={editarReceitaClt}
                    excluir={(id) =>
                      setReceitasClt(
                        receitasClt.filter((item) => item.id !== id),
                      )
                    }
                  />

                  <TabelaLancamentos
                    titulo="MEI"
                    cor="green"
                    itens={receitasMei}
                    novoNome={novaMeiNome}
                    novoValor={novaMeiValor}
                    setNovoNome={setNovaMeiNome}
                    setNovoValor={setNovaMeiValor}
                    adicionar={adicionarReceitaMei}
                    editar={editarReceitaMei}
                    excluir={(id) =>
                      setReceitasMei(
                        receitasMei.filter((item) => item.id !== id),
                      )
                    }
                  />
                </div>

                <div className="grid grid-cols-3 gap-4 mt-8">
                  <Card
                    titulo="Total CLT"
                    valor={formatar(totalReceitasClt)}
                    cor="text-green-400"
                  />
                  <Card
                    titulo="Total MEI"
                    valor={formatar(totalReceitasMei)}
                    cor="text-emerald-400"
                  />
                  <Card
                    titulo="Receita total"
                    valor={formatar(totalReceitas)}
                    cor="text-green-500"
                  />
                </div>
              </>
            )}

            {subAbaMensal === "despesas" && (
              <>
                <TabelaLancamentos
                  titulo="Despesas"
                  cor="red"
                  itens={despesasMensais}
                  novoNome={novaDespesaNome}
                  novoValor={novaDespesaValor}
                  setNovoNome={setNovaDespesaNome}
                  setNovoValor={setNovaDespesaValor}
                  adicionar={adicionarDespesaMensal}
                  editar={editarDespesaMensal}
                  excluir={(id) =>
                    setDespesasMensais(
                      despesasMensais.filter((item) => item.id !== id),
                    )
                  }
                />

                <div className="grid grid-cols-4 gap-4 mt-8">
                  <Card
                    titulo="Despesas mensais"
                    valor={formatar(totalDespesasMensais)}
                    cor="text-red-500"
                  />
                  <Card
                    titulo="Cartões"
                    valor={formatar(totalComprasCartao)}
                    cor="text-red-400"
                  />
                  <Card
                    titulo="Assinaturas"
                    valor={formatar(totalAssinaturas)}
                    cor="text-orange-400"
                  />
                  <Card
                    titulo="Total despesas"
                    valor={formatar(totalDespesas)}
                    cor="text-red-500"
                  />
                </div>
              </>
            )}
          </>
        )}

        {aba === "cartoes" && (
          <>
            <h2 className="text-4xl font-bold mb-4">Cartões</h2>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl mb-6">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-2xl font-bold">Cartões de Rennan</h3>
                  <p className="text-zinc-400 text-sm mt-1">
                    Os totais consideram somente Rennan + assinaturas fixas do cartão.
                  </p>
                </div>

                <div className="text-right">
                  <p className="text-zinc-400 text-sm">Total dos cartões</p>
                  <p className="text-2xl font-bold text-red-500">
                    {formatar(totalCartoesFiltrado)}
                  </p>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4">
              {gastosPorCartao.map((cartao) => (
                <div
                  key={cartao.nome}
                  className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl"
                >
                  <h3 className="text-2xl font-bold mb-4">{cartao.nome}</h3>
                  <p className="text-zinc-400">Total do cartão</p>
                  <p className="text-3xl font-bold text-red-500">
                    {formatar(cartao.gasto)}
                  </p>

                  <p className="text-zinc-400 mt-4">Limite</p>
                  <input
                    className="w-full bg-zinc-800 p-3 rounded mt-2 outline-none"
                    value={cartao.limite}
                    type="number"
                    onChange={(e) => alterarLimite(cartao.nome, e.target.value)}
                  />

                  <p className="text-sm text-zinc-400 mt-3">
                    Fecha dia {faturasCartoes.find((item) => item.nome === cartao.nome)?.fechamento || "-"} • Vence dia {faturasCartoes.find((item) => item.nome === cartao.nome)?.vencimento || "-"}
                  </p>

                  <div className="mt-4 flex gap-2">
                    <button
                      type="button"
                      onClick={() => setFaturaDetalhesAberta(cartao.nome)}
                      className="text-xs bg-blue-700 hover:bg-blue-600 px-3 py-2 rounded text-white"
                    >
                      Ver fatura
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        setCartaoVencimentoAberto(
                          cartaoVencimentoAberto === cartao.nome ? null : cartao.nome,
                        )
                      }
                      className="text-xs bg-zinc-800 hover:bg-zinc-700 px-3 py-2 rounded text-zinc-300"
                    >
                      Vencimento
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        setCartaoAssinaturasAberto(
                          cartaoAssinaturasAberto === cartao.nome ? null : cartao.nome,
                        )
                      }
                      className="text-xs bg-zinc-800 hover:bg-zinc-700 px-3 py-2 rounded text-zinc-300"
                    >
                      Assinaturas
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        setCartaoAvulsoAberto(
                          cartaoAvulsoAberto === cartao.nome ? null : cartao.nome,
                        )
                      }
                      className="text-xs bg-zinc-800 hover:bg-zinc-700 px-3 py-2 rounded text-zinc-300"
                    >
                      Avulso
                    </button>
                  </div>

                  {cartaoVencimentoAberto === cartao.nome && (
                    <div className="mt-4 grid grid-cols-2 gap-3 border-t border-zinc-800 pt-4">
                      <div>
                        <p className="mb-1 text-xs text-zinc-500">Fechamento</p>
                        <input
                          className="w-full rounded bg-zinc-800 p-3 outline-none"
                          type="number"
                          min="1"
                          max="31"
                          value={faturasCartoes.find((item) => item.nome === cartao.nome)?.fechamento || 1}
                          onChange={(e) =>
                            alterarFaturaCartao(cartao.nome, "fechamento", e.target.value)
                          }
                        />
                      </div>

                      <div>
                        <p className="mb-1 text-xs text-zinc-500">Vencimento</p>
                        <input
                          className="w-full rounded bg-zinc-800 p-3 outline-none"
                          type="number"
                          min="1"
                          max="31"
                          value={faturasCartoes.find((item) => item.nome === cartao.nome)?.vencimento || 1}
                          onChange={(e) =>
                            alterarFaturaCartao(cartao.nome, "vencimento", e.target.value)
                          }
                        />
                      </div>
                    </div>
                  )}

                  {cartaoAssinaturasAberto === cartao.nome && (
                    <div className="mt-4 border-t border-zinc-800 pt-4">
                      <div className="flex items-center justify-between mb-3">
                        <p className="text-sm font-bold text-zinc-300">
                          Assinaturas fixas
                        </p>
                        <p className="text-xs text-zinc-500">
                          {formatar(
                            assinaturas
                              .filter((item) => item.cartao === cartao.nome && item.pessoa === "Rennan")
                              .reduce((soma, item) => soma + item.valor, 0),
                          )}
                        </p>
                      </div>

                      <div className="grid grid-cols-[1fr_120px_140px_90px] gap-2 mb-4">
                        <input
                          className="bg-zinc-800 p-2 rounded outline-none text-sm"
                          placeholder="Assinatura"
                          value={assinaturaCartao === cartao.nome ? assinaturaDescricao : ""}
                          onChange={(e) => {
                            setAssinaturaCartao(cartao.nome);
                            setAssinaturaDescricao(e.target.value);
                          }}
                        />

                        <input
                          className="bg-zinc-800 p-2 rounded outline-none text-sm"
                          placeholder="Valor"
                          value={assinaturaCartao === cartao.nome ? assinaturaValor : ""}
                          onChange={(e) => {
                            setAssinaturaCartao(cartao.nome);
                            setAssinaturaValor(e.target.value);
                          }}
                        />

                        <select
                          className="bg-zinc-800 p-2 rounded outline-none text-sm"
                          value={assinaturaCartao === cartao.nome ? assinaturaPessoa : "Rennan"}
                          onChange={(e) => {
                            setAssinaturaCartao(cartao.nome);
                            setAssinaturaPessoa(e.target.value);
                          }}
                        >
                          {pessoas.map((nome) => (
                            <option key={nome}>{nome}</option>
                          ))}
                        </select>

                        <button
                          type="button"
                          onClick={() => salvarAssinatura(cartao.nome)}
                          className="bg-green-600 hover:bg-green-700 rounded font-bold text-sm"
                        >
                          {assinaturaEditandoId && assinaturaCartao === cartao.nome ? "Salvar" : "+ Add"}
                        </button>
                      </div>

                      {assinaturas.filter((item) => item.cartao === cartao.nome).length === 0 ? (
                        <p className="text-sm text-zinc-500">
                          Nenhuma assinatura fixa neste cartão.
                        </p>
                      ) : (
                        <div className="space-y-3">
                          {assinaturas
                            .filter((item) => item.cartao === cartao.nome)
                            .map((item) => (
                              <div
                                key={item.id}
                                className="flex items-center justify-between gap-3 border-b border-zinc-800 pb-3"
                              >
                                <div>
                                  <p className="font-medium">{item.descricao}</p>
                                  <p className="text-xs text-zinc-500">{item.pessoa || "Rennan"} • Todo mês</p>
                                </div>

                                <div className="flex items-center gap-3">
                                  <span className="text-orange-400 font-bold">
                                    {formatar(item.valor)}
                                  </span>

                                  <button
                                    type="button"
                                    onClick={() => abrirEditarAssinatura(item)}
                                    className="text-blue-400 hover:text-blue-300 text-xs"
                                  >
                                    Editar
                                  </button>

                                  <button
                                    type="button"
                                    onClick={() => excluirAssinatura(item.id)}
                                    className="text-red-400 hover:text-red-300 text-xs"
                                  >
                                    Excluir
                                  </button>
                                </div>
                              </div>
                            ))}
                        </div>
                      )}
                    </div>
                  )}

                  {cartaoAvulsoAberto === cartao.nome && (
                    <div className="mt-4 border-t border-zinc-800 pt-4">
                      <div className="flex items-center justify-between mb-3">
                        <p className="text-sm font-bold text-zinc-300">
                          Gastos avulsos de Rennan
                        </p>
                        <p className="text-xs text-zinc-500">
                          {formatar(
                            compras
                              .filter(
                                (item) =>
                                  item.cartao === cartao.nome &&
                                  item.pessoa === "Rennan" &&
                                  (item.tipo || (item.parcelas > 1 ? "parcela" : "avulso")) === "avulso",
                              )
                              .reduce((soma, item) => soma + item.valor, 0),
                          )}
                        </p>
                      </div>

                      {compras.filter(
                        (item) =>
                          item.cartao === cartao.nome &&
                          item.pessoa === "Rennan" &&
                          (item.tipo || (item.parcelas > 1 ? "parcela" : "avulso")) === "avulso",
                      ).length === 0 ? (
                        <p className="text-sm text-zinc-500">
                          Nenhum gasto avulso cadastrado neste cartão.
                        </p>
                      ) : (
                        <div className="space-y-3">
                          {compras
                            .filter(
                              (item) =>
                                item.cartao === cartao.nome &&
                                item.pessoa === "Rennan" &&
                                (item.tipo || (item.parcelas > 1 ? "parcela" : "avulso")) === "avulso",
                            )
                            .map((item) => (
                              <div
                                key={item.id}
                                className="flex items-center justify-between gap-3 border-b border-zinc-800 pb-3"
                              >
                                <div>
                                  <p className="font-medium">{item.descricao}</p>
                                  <p className="text-xs text-zinc-500">
                                    {item.pessoa} • {item.data}
                                  </p>
                                </div>

                                <div className="flex items-center gap-3">
                                  <span className="text-red-400 font-bold">
                                    {formatar(item.valor)}
                                  </span>

                                  <button
                                    type="button"
                                    onClick={() => abrirModalEditarCompra(item)}
                                    className="text-blue-400 hover:text-blue-300 text-xs"
                                  >
                                    Editar
                                  </button>

                                  <button
                                    type="button"
                                    onClick={() => excluirCompra(item.id)}
                                    className="text-red-400 hover:text-red-300 text-xs"
                                  >
                                    Excluir
                                  </button>
                                </div>
                              </div>
                            ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </>
        )}

        {aba === "parcelas" && (
          <>
            <div className="flex items-center justify-between mb-8">
              <div>
                <h2 className="text-4xl font-bold">Parcelas por Cartão</h2>
                <p className="text-zinc-400 mt-2">
                  Cadastre, edite, exclua e filtre as parcelas por pessoa.
                </p>
              </div>

              <div className="text-right">
                <p className="text-zinc-400 text-sm">
                  Total de parcelas filtradas
                </p>
                <p className="text-3xl font-bold text-yellow-400">
                  {formatar(totalParcelasSelecionadas)}
                </p>
              </div>
            </div>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl mb-6">
              <h3 className="text-2xl font-bold mb-4">Filtro por pessoa</h3>
              <FiltroPessoasCartoes
                pessoas={pessoas}
                selecionadas={pessoasSelecionadasParcelas}
                alternar={alternarPessoaParcelas}
                selecionarTodas={selecionarTodasPessoasParcelas}
              />
            </div>

            <div className="grid grid-cols-1 gap-6">
              {parcelasPorCartao.every(
                (cartao) => cartao.itens.length === 0,
              ) && (
                <p className="text-zinc-500">
                  Nenhuma parcela cadastrada para o filtro selecionado.
                </p>
              )}

              {parcelasPorCartao.map((cartao) => (
                <div
                  key={cartao.nome}
                  className="bg-zinc-900 border border-zinc-800 rounded-xl overflow-hidden"
                >
                  <div className="flex items-center justify-between bg-zinc-800 p-4">
                    <h3 className="text-2xl font-bold">💳 {cartao.nome}</h3>
                    <p className="text-yellow-400 font-bold">
                      Total de parcelas: {formatar(cartao.totalParcelas)}
                    </p>
                  </div>

                  {cartao.itens.length === 0 ? (
                    <p className="p-4 text-zinc-500">
                      Sem parcelas neste cartão.
                    </p>
                  ) : (
                    <table className="w-full">
                      <thead className="bg-zinc-800/60">
                        <tr>
                          <th className="p-4 text-left">Compra</th>
                          <th className="p-4 text-left">Pessoa</th>
                          <th className="p-4 text-left">Valor da parcela</th>
                          <th className="p-4 text-left">Parcela</th>
                          <th className="p-4 text-left">Parcelas restantes</th>
                          <th className="p-4 text-left">Total de parcelas</th>
                          <th className="p-4 text-left">Ação</th>
                        </tr>
                      </thead>

                      <tbody>
                        {cartao.itens.map((item) => {
                          const atual = item.parcelaAtual || 1;
                          const restantes = Math.max(
                            item.parcelas - atual + 1,
                            0,
                          );

                          return (
                            <tr
                              key={item.id}
                              className="border-t border-zinc-800"
                            >
                              <td className="p-4">{item.descricao}</td>
                              <td className="p-4">{item.pessoa}</td>
                              <td className="p-4 text-yellow-400 font-bold">
                                {formatar(item.valor)}
                              </td>
                              <td className="p-4">
                                {atual}/{item.parcelas}
                              </td>
                              <td className="p-4">{restantes}</td>
                              <td className="p-4 text-yellow-400 font-bold">
                                {formatar(item.valor * restantes)}
                              </td>
                              <td className="p-4">
                                <div className="flex gap-3">
                                  <button
                                    type="button"
                                    onClick={() => abrirModalEditarCompra(item)}
                                    className="text-blue-400 hover:text-blue-300"
                                  >
                                    Editar
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => excluirCompra(item.id)}
                                    className="text-red-400 hover:text-red-300"
                                  >
                                    Excluir
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}

                        <tr className="border-t border-zinc-700 bg-zinc-800 font-bold">
                          <td className="p-4" colSpan={5}>
                            Total de parcelas do cartão
                          </td>
                          <td className="p-4 text-yellow-400">
                            {formatar(cartao.totalParcelas)}
                          </td>
                          <td className="p-4" />
                        </tr>
                      </tbody>
                    </table>
                  )}
                </div>
              ))}
            </div>
          </>
        )}

        {aba === "pessoas" && (
          <>
            <h2 className="text-4xl font-bold mb-8">Pessoas</h2>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl mb-6">
              <div className="flex gap-3">
                <input
                  className="flex-1 bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Nome da pessoa"
                  value={novaPessoa}
                  onChange={(e) => setNovaPessoa(e.target.value)}
                />

                <button
                  type="button"
                  onClick={adicionarPessoa}
                  className="bg-green-600 hover:bg-green-700 px-6 rounded font-bold"
                >
                  Adicionar
                </button>
              </div>
            </div>

            <div className="grid grid-cols-4 gap-4">
              {pessoas.map((nome) => {
                const total = totalDaPessoa(nome);

                return (
                  <div
                    key={nome}
                    className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl"
                  >
                    <p className="text-zinc-400">{nome}</p>
                    <p className="text-2xl font-bold mb-4">{formatar(total)}</p>

                    <div className="flex gap-3">
                      <button
                        type="button"
                        onClick={() => setPessoaDetalhes(nome)}
                        className="text-blue-400 hover:text-blue-300 text-sm"
                      >
                        Detalhes
                      </button>

                      <button
                        type="button"
                        onClick={() => excluirPessoa(nome)}
                        className="text-red-400 hover:text-red-300 text-sm"
                      >
                        Excluir pessoa
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </>
        )}

        {aba === "historico" && (
          <>
            <div className="flex items-center justify-between mb-8">
              <div>
                <h2 className="text-4xl font-bold">Histórico Mensal</h2>
                <p className="text-zinc-400 mt-2">
                  O sistema salva o mês automaticamente quando virar o mês.
                </p>
              </div>

              <button
                type="button"
                onClick={salvarHistoricoMesAtual}
                className="bg-green-600 hover:bg-green-700 px-5 py-3 rounded-lg font-bold"
              >
                Salvar mês atual
              </button>
            </div>

            {historicoMensal.length === 0 ? (
              <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl">
                <p className="text-zinc-400">
                  Nenhum mês salvo ainda. Clique em "Salvar mês atual" para criar o primeiro histórico.
                </p>
              </div>
            ) : (
              <div className="grid grid-cols-1 gap-6">
                {[...historicoMensal]
                  .sort((a, b) => b.mes.localeCompare(a.mes))
                  .map((item) => (
                    <div
                      key={item.mes}
                      className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl"
                    >
                      <div className="flex items-center justify-between mb-5">
                        <div>
                          <h3 className="text-2xl font-bold capitalize">
                            {nomeMesHistorico(item.mes)}
                          </h3>
                          <p className="text-zinc-500 text-sm">
                            Salvo em {item.criadoEm}
                          </p>
                        </div>

                        <button
                          type="button"
                          onClick={() => excluirHistorico(item.mes)}
                          className="text-red-400 hover:text-red-300 text-sm"
                        >
                          Excluir histórico
                        </button>
                      </div>

                      <div className="grid grid-cols-4 gap-4">
                        <Card
                          titulo="Receitas"
                          valor={formatar(item.receitas)}
                          cor="text-green-500"
                        />
                        <Card
                          titulo="Despesas totais"
                          valor={formatar(item.despesasTotais)}
                          cor="text-red-500"
                        />
                        <Card
                          titulo="Sobra"
                          valor={formatar(item.sobra)}
                          cor={item.sobra >= 0 ? "text-blue-400" : "text-red-500"}
                        />
                        <Card
                          titulo="Parcelas futuras"
                          valor={formatar(item.parcelasFuturas)}
                          cor="text-yellow-400"
                        />
                      </div>

                      <div className="grid grid-cols-4 gap-4 mt-4">
                        <Card
                          titulo="Despesas mensais"
                          valor={formatar(item.despesasMensais)}
                          cor="text-red-400"
                        />
                        <Card
                          titulo="Cartões"
                          valor={formatar(item.cartoes)}
                          cor="text-red-400"
                        />
                        <Card
                          titulo="Assinaturas"
                          valor={formatar(item.assinaturas)}
                          cor="text-orange-400"
                        />
                        <Card
                          titulo="Resultado"
                          valor={formatar(item.sobra)}
                          cor={item.sobra >= 0 ? "text-blue-400" : "text-red-500"}
                        />
                      </div>
                    </div>
                  ))}
              </div>
            )}
          </>
        )}

        {aba === "config" && (
          <>
            <h2 className="text-4xl font-bold mb-8">Configurações</h2>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl max-w-4xl mb-6">
              <h3 className="text-2xl font-bold mb-3">Gerenciar cartões</h3>
              <p className="text-zinc-400 mb-4">
                Cadastre novos cartões ou exclua cartões que não usa mais.
              </p>

              <div className="grid grid-cols-[1fr_180px_150px] gap-3 mb-5">
                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Nome do cartão"
                  value={novoCartaoNome}
                  onChange={(e) => setNovoCartaoNome(e.target.value)}
                />

                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Limite"
                  value={novoCartaoLimite}
                  onChange={(e) => setNovoCartaoLimite(e.target.value)}
                />

                <button
                  type="button"
                  onClick={adicionarCartao}
                  className="bg-green-600 hover:bg-green-700 rounded font-bold"
                >
                  + Adicionar
                </button>
              </div>

              <div className="space-y-3">
                {cartoes.map((cartaoItem) => {
                  const temMovimento =
                    compras.some((item) => item.cartao === cartaoItem.nome) ||
                    assinaturas.some((item) => item.cartao === cartaoItem.nome);

                  return (
                    <div
                      key={cartaoItem.nome}
                      className="flex items-center justify-between bg-zinc-800/60 border border-zinc-700 p-3 rounded-lg"
                    >
                      <div>
                        <p className="font-bold">{cartaoItem.nome}</p>
                        <p className="text-xs text-zinc-500">
                          Limite: {formatar(cartaoItem.limite)}
                          {temMovimento ? " • possui movimentação" : ""}
                        </p>
                      </div>

                      <button
                        type="button"
                        onClick={() => excluirCartao(cartaoItem.nome)}
                        className="text-sm bg-red-950/60 hover:bg-red-900 px-4 py-2 rounded text-red-300"
                      >
                        Excluir
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl max-w-4xl mb-6">
              <h3 className="text-2xl font-bold mb-3">Cadastrar gasto avulso</h3>
              <p className="text-zinc-400 mb-4">
                Use para compras feitas uma única vez, sem parcelamento.
              </p>

              <div className="grid grid-cols-5 gap-3">
                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Compra"
                  value={descricao}
                  onChange={(e) => setDescricao(e.target.value)}
                />

                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Valor"
                  value={valor}
                  onChange={(e) => setValor(e.target.value)}
                />

                <select
                  className="bg-zinc-800 p-3 rounded outline-none"
                  value={cartao}
                  onChange={(e) => setCartao(e.target.value)}
                >
                  {cartoes.map((item) => (
                    <option key={item.nome}>{item.nome}</option>
                  ))}
                </select>

                <select
                  className="bg-zinc-800 p-3 rounded outline-none"
                  value={pessoa}
                  onChange={(e) => setPessoa(e.target.value)}
                >
                  {pessoas.map((item) => (
                    <option key={item}>{item}</option>
                  ))}
                </select>

                <button
                  type="button"
                  onClick={salvarGastoAvulso}
                  className="bg-green-600 hover:bg-green-700 rounded font-bold"
                >
                  + Adicionar
                </button>
              </div>
            </div>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl max-w-4xl mb-6">
              <h3 className="text-2xl font-bold mb-3">Cadastrar parcela</h3>
              <p className="text-zinc-400 mb-4">
                Use para compras parceladas. Depois elas aparecerão na aba Parcelas.
              </p>

              <div className="grid grid-cols-6 gap-3">
                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Compra"
                  value={descricao}
                  onChange={(e) => setDescricao(e.target.value)}
                />

                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Valor da parcela"
                  value={valor}
                  onChange={(e) => setValor(e.target.value)}
                />

                <select
                  className="bg-zinc-800 p-3 rounded outline-none"
                  value={cartao}
                  onChange={(e) => setCartao(e.target.value)}
                >
                  {cartoes.map((item) => (
                    <option key={item.nome}>{item.nome}</option>
                  ))}
                </select>

                <select
                  className="bg-zinc-800 p-3 rounded outline-none"
                  value={pessoa}
                  onChange={(e) => setPessoa(e.target.value)}
                >
                  {pessoas.map((item) => (
                    <option key={item}>{item}</option>
                  ))}
                </select>

                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Parcela atual"
                  type="number"
                  min="1"
                  value={parcelaAtual}
                  onChange={(e) => setParcelaAtual(e.target.value)}
                />

                <input
                  className="bg-zinc-800 p-3 rounded outline-none"
                  placeholder="Total parcelas"
                  type="number"
                  min="1"
                  value={parcelas}
                  onChange={(e) => setParcelas(e.target.value)}
                />
              </div>

              <button
                type="button"
                onClick={salvarCompra}
                className="mt-4 bg-green-600 hover:bg-green-700 px-6 py-3 rounded-lg font-bold"
              >
                + Cadastrar parcela
              </button>
            </div>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl max-w-2xl mb-6">
              <h3 className="text-2xl font-bold mb-3">Backup dos dados</h3>
              <p className="text-zinc-400 mb-4">
                Baixe um arquivo com todos os dados do sistema para não perder nada caso limpe o navegador ou troque de computador.
              </p>

              <div className="flex flex-col gap-3">
                <button
                  type="button"
                  onClick={baixarBackup}
                  className="bg-blue-600 hover:bg-blue-700 px-5 py-3 rounded-lg font-bold"
                >
                  Baixar backup
                </button>

                <label className="block text-sm text-zinc-400">
                  Restaurar backup
                  <input
                    type="file"
                    accept=".json"
                    onChange={(e) => restaurarBackup(e.target.files?.[0])}
                    className="mt-3 block w-full text-sm text-zinc-300 file:mr-4 file:rounded-lg file:border-0 file:bg-orange-600 file:px-4 file:py-3 file:font-bold file:text-white hover:file:bg-orange-700"
                  />
                </label>
              </div>

              <div className="mt-5 text-sm text-zinc-400 space-y-1">
                <p>✅ Salva compras, parcelas e avulsos</p>
                <p>✅ Salva pessoas, cartões e limites</p>
                <p>✅ Salva assinaturas, receitas, despesas e histórico</p>
              </div>
            </div>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl max-w-2xl mb-6">
              <h3 className="text-2xl font-bold mb-3">Aplicativo no celular</h3>
              <p className="text-zinc-400 mb-4">
                Com os arquivos de PWA instalados, abra o site no celular e use a opção do navegador para adicionar à tela inicial.
              </p>

              <div className="text-sm text-zinc-400 space-y-1">
                <p>✅ Android/Chrome: menu ⋮ → Adicionar à tela inicial</p>
                <p>✅ iPhone/Safari: compartilhar → Adicionar à Tela de Início</p>
                <p>✅ O sistema também registra um service worker básico para abrir como app.</p>
              </div>
            </div>

            <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl max-w-2xl">
              <h3 className="text-2xl font-bold mb-3">
                Importar planilha Excel
              </h3>
              <p className="text-zinc-400 mb-4">
                Use a sua planilha de controle financeiro para preencher
                automaticamente CLT, MEI, despesas mensais e compras dos
                cartões.
              </p>

              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => importarPlanilha(e.target.files?.[0])}
                className="block w-full text-sm text-zinc-300 file:mr-4 file:rounded-lg file:border-0 file:bg-green-600 file:px-4 file:py-3 file:font-bold file:text-white hover:file:bg-green-700"
              />

              <div className="mt-5 text-sm text-zinc-400 space-y-1">
                <p>✅ Aba Mensal: importa CLT, MEI e PIX/Débito</p>
                <p>✅ Abas dos cartões: importa compras por pessoa</p>
                <p>
                  ✅ Parcelas: identifica automaticamente textos como (10/12)
                </p>
              </div>
            </div>
          </>
        )}
      </section>

      {modalAberto && (
        <div className="fixed inset-0 bg-black/70 flex items-center justify-center z-50">
          <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl w-[420px]">
            <h2 className="text-2xl font-bold mb-4">
              {compraEditandoId ? "Editar Parcela" : "Nova Parcela"}
            </h2>

            <input
              className="w-full bg-zinc-800 p-3 rounded mb-3 outline-none"
              placeholder="Descrição"
              value={descricao}
              onChange={(e) => setDescricao(e.target.value)}
            />
            <input
              className="w-full bg-zinc-800 p-3 rounded mb-3 outline-none"
              placeholder="Valor"
              value={valor}
              onChange={(e) => setValor(e.target.value)}
            />

            <select
              className="w-full bg-zinc-800 p-3 rounded mb-3 outline-none"
              value={cartao}
              onChange={(e) => setCartao(e.target.value)}
            >
              {cartoes.map((item) => (
                <option key={item.nome}>{item.nome}</option>
              ))}
            </select>

            <select
              className="w-full bg-zinc-800 p-3 rounded mb-3 outline-none"
              value={pessoa}
              onChange={(e) => setPessoa(e.target.value)}
            >
              {pessoas.map((item) => (
                <option key={item}>{item}</option>
              ))}
            </select>

            <div className="grid grid-cols-2 gap-3">
              <input
                className="w-full bg-zinc-800 p-3 rounded mb-5 outline-none"
                placeholder="Parcela atual"
                type="number"
                min="1"
                value={parcelaAtual}
                onChange={(e) => setParcelaAtual(e.target.value)}
              />
              <input
                className="w-full bg-zinc-800 p-3 rounded mb-5 outline-none"
                placeholder="Total parcelas"
                type="number"
                min="1"
                value={parcelas}
                onChange={(e) => setParcelas(e.target.value)}
              />
            </div>

            <div className="flex gap-3">
              <button
                type="button"
                onClick={salvarCompra}
                className="flex-1 bg-green-600 hover:bg-green-700 p-3 rounded font-bold"
              >
                {compraEditandoId ? "Salvar edição" : "Salvar"}
              </button>

              <button
                type="button"
                onClick={() => {
                  limparFormularioCompra();
                  setModalAberto(false);
                }}
                className="flex-1 bg-zinc-700 hover:bg-zinc-600 p-3 rounded font-bold"
              >
                Cancelar
              </button>
            </div>
          </div>
        </div>
      )}

      {pessoaDetalhes && (
        <div className="fixed inset-0 bg-black/70 flex items-center justify-center z-50">
          <div className="bg-zinc-900 border border-zinc-800 rounded-xl w-[980px] max-h-[80vh] overflow-hidden">
            <div className="flex items-center justify-between p-6 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl font-bold">
                  Detalhes de {pessoaDetalhes}
                </h2>
                <p className="text-zinc-400">
                  Total atual: {formatar(totalDaPessoa(pessoaDetalhes))} • Total
                  de parcelas: {formatar(totalParcelasDaPessoa(pessoaDetalhes))} • Assinaturas: {formatar(totalAssinaturasDaPessoa(pessoaDetalhes))}
                </p>
              </div>

              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => baixarExcelPessoa(pessoaDetalhes)}
                  className="bg-green-600 hover:bg-green-700 px-4 py-2 rounded-lg font-bold"
                >
                  Baixar Excel
                </button>

                <button
                  type="button"
                  onClick={() => baixarPdfPessoa(pessoaDetalhes)}
                  className="bg-red-600 hover:bg-red-700 px-4 py-2 rounded-lg font-bold"
                >
                  Baixar PDF
                </button>

                <button
                  type="button"
                  onClick={() => setPessoaDetalhes(null)}
                  className="bg-zinc-800 hover:bg-zinc-700 px-4 py-2 rounded-lg font-bold"
                >
                  Fechar
                </button>
              </div>
            </div>

            <div className="overflow-auto max-h-[60vh]">
              {comprasDaPessoa(pessoaDetalhes).length === 0 && assinaturasDaPessoa(pessoaDetalhes).length === 0 ? (
                <p className="p-6 text-zinc-400">
                  Essa pessoa não possui compras ou assinaturas cadastradas.
                </p>
              ) : (
                <table className="w-full">
                  <thead className="bg-zinc-800">
                    <tr>
                      <th className="p-4 text-left">Compra</th>
                      <th className="p-4 text-left">Valor da parcela</th>
                      <th className="p-4 text-left">Parcela</th>
                      <th className="p-4 text-left">Parcelas restantes</th>
                      <th className="p-4 text-left">Total de parcelas</th>
                      <th className="p-4 text-left">Data</th>
                      <th className="p-4 text-left">Ação</th>
                    </tr>
                  </thead>

                  <tbody>
                    {comprasDaPessoa(pessoaDetalhes).map((item) => {
                      const atual = item.parcelaAtual || 1;
                      const restantes = Math.max(item.parcelas - atual + 1, 0);

                      return (
                        <tr key={item.id} className="border-t border-zinc-800">
                          <td className="p-4">{item.descricao}</td>
                          <td className="p-4 text-red-500 font-bold">
                            {formatar(item.valor)}
                          </td>
                          <td className="p-4">
                            {atual}/{item.parcelas}
                          </td>
                          <td className="p-4">{restantes}</td>
                          <td className="p-4 text-yellow-400 font-bold">
                            {formatar(item.valor * restantes)}
                          </td>
                          <td className="p-4">{item.data}</td>
                          <td className="p-4">
                            <div className="flex gap-3">
                              <button
                                type="button"
                                onClick={() => abrirModalEditarCompra(item)}
                                className="text-blue-400 hover:text-blue-300"
                              >
                                Editar
                              </button>

                              <button
                                type="button"
                                onClick={() => excluirCompra(item.id)}
                                className="text-red-400 hover:text-red-300"
                              >
                                Excluir
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}

                    {assinaturasDaPessoa(pessoaDetalhes).map((item) => (
                      <tr key={`assinatura-${item.id}`} className="border-t border-zinc-800">
                        <td className="p-4">{item.descricao}</td>
                        <td className="p-4 text-orange-400 font-bold">
                          {formatar(item.valor)}
                        </td>
                        <td className="p-4">fixa</td>
                        <td className="p-4">1</td>
                        <td className="p-4 text-yellow-400 font-bold">
                          {formatar(item.valor)}
                        </td>
                        <td className="p-4">Todo mês</td>
                        <td className="p-4">
                          <div className="flex gap-3">
                            <button
                              type="button"
                              onClick={() => {
                                abrirEditarAssinatura(item);
                                setAba("cartoes");
                                setPessoaDetalhes(null);
                                setCartaoAssinaturasAberto(item.cartao);
                              }}
                              className="text-blue-400 hover:text-blue-300"
                            >
                              Editar
                            </button>

                            <button
                              type="button"
                              onClick={() => excluirAssinatura(item.id)}
                              className="text-red-400 hover:text-red-300"
                            >
                              Excluir
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}

                    <tr className="border-t border-zinc-700 bg-zinc-800 font-bold">
                      <td className="p-4" colSpan={4}>
                        Total de parcelas da pessoa
                      </td>
                      <td className="p-4 text-yellow-400">
                        {formatar(totalParcelasDaPessoa(pessoaDetalhes))}
                      </td>
                      <td className="p-4" />
                      <td className="p-4" />
                    </tr>
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}

      {faturaDetalhesAberta && (() => {
        const cartaoSelecionado = gastosPorCartao.find(
          (item) => item.nome === faturaDetalhesAberta,
        );
        const fatura = faturasResumo.find(
          (item) => item.nome === faturaDetalhesAberta,
        );
        const comprasDoCartao = comprasFiltradasCartoes.filter(
          (item) => item.cartao === faturaDetalhesAberta,
        );
        const assinaturasDoCartao = assinaturas.filter(
          (item) => item.cartao === faturaDetalhesAberta && item.pessoa === "Rennan",
        );
        const limite = cartaoSelecionado?.limite || 0;
        const utilizado = cartaoSelecionado?.gasto || 0;
        const disponivel = Math.max(limite - utilizado, 0);
        const uso = cartaoSelecionado?.uso || 0;

        return (
          <div className="fixed inset-0 bg-black/70 flex items-center justify-center z-50">
            <div className="bg-zinc-900 border border-zinc-800 rounded-xl w-[880px] max-h-[85vh] overflow-hidden">
              <div className="flex items-center justify-between p-6 border-b border-zinc-800">
                <div>
                  <h2 className="text-2xl font-bold">
                    💳 Fatura {faturaDetalhesAberta}
                  </h2>
                  <p className="text-zinc-400 mt-1">
                    Fecha dia {fatura?.fechamento || "-"} • Vence dia {fatura?.vencimento || "-"}
                    {fatura ? ` • ${formatarDataCurta(fatura.dataVencimento)} • ${fatura.dias} dias` : ""}
                  </p>
                </div>

                <button
                  type="button"
                  onClick={() => setFaturaDetalhesAberta(null)}
                  className="bg-zinc-800 hover:bg-zinc-700 px-4 py-2 rounded-lg font-bold"
                >
                  Fechar
                </button>
              </div>

              <div className="p-6 overflow-auto max-h-[70vh]">
                <div className="grid grid-cols-3 gap-4 mb-6">
                  <Card titulo="Limite" valor={formatar(limite)} cor="text-blue-400" />
                  <Card titulo="Utilizado" valor={formatar(utilizado)} cor="text-red-400" />
                  <Card titulo="Disponível" valor={formatar(disponivel)} cor="text-green-400" />
                </div>

                <div className="mb-6">
                  <div className="w-full bg-zinc-800 rounded-full h-3">
                    <div
                      className="bg-green-500 h-3 rounded-full"
                      style={{ width: `${uso}%` }}
                    />
                  </div>
                  <p className="text-sm text-zinc-400 mt-2">Uso: {uso.toFixed(0)}%</p>
                </div>

                <div className="bg-zinc-950 border border-zinc-800 rounded-xl overflow-hidden">
                  <table className="w-full">
                    <thead className="bg-zinc-800">
                      <tr>
                        <th className="p-4 text-left">Descrição</th>
                        <th className="p-4 text-left">Tipo</th>
                        <th className="p-4 text-left">Parcela</th>
                        <th className="p-4 text-left">Valor</th>
                      </tr>
                    </thead>

                    <tbody>
                      {comprasDoCartao.length === 0 && assinaturasDoCartao.length === 0 && (
                        <tr>
                          <td className="p-4 text-zinc-400" colSpan={4}>
                            Nenhum item cadastrado nesta fatura.
                          </td>
                        </tr>
                      )}

                      {comprasDoCartao.map((item) => (
                        <tr key={item.id} className="border-t border-zinc-800">
                          <td className="p-4">{item.descricao}</td>
                          <td className="p-4 capitalize">{tipoDaCompra(item)}</td>
                          <td className="p-4">
                            {item.parcelaAtual || 1}/{item.parcelas}
                          </td>
                          <td className="p-4 font-bold text-red-400">
                            {formatar(item.valor)}
                          </td>
                        </tr>
                      ))}

                      {assinaturasDoCartao.map((item) => (
                        <tr key={`assinatura-fatura-${item.id}`} className="border-t border-zinc-800">
                          <td className="p-4">{item.descricao}</td>
                          <td className="p-4">assinatura</td>
                          <td className="p-4">fixa</td>
                          <td className="p-4 font-bold text-orange-400">
                            {formatar(item.valor)}
                          </td>
                        </tr>
                      ))}

                      <tr className="border-t border-zinc-700 bg-zinc-800 font-bold">
                        <td className="p-4" colSpan={3}>Total da fatura</td>
                        <td className="p-4 text-red-400">{formatar(utilizado)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

    </main>
  );
}

function FiltroPessoasCartoes({
  pessoas,
  selecionadas,
  alternar,
  selecionarTodas,
}: {
  pessoas: string[];
  selecionadas: string[];
  alternar: (nome: string) => void;
  selecionarTodas: () => void;
}) {
  const todosSelecionados =
    pessoas.length > 0 && selecionadas.length === pessoas.length;

  return (
    <div className="flex flex-wrap gap-3 mb-6">
      <button
        type="button"
        onClick={selecionarTodas}
        className={`px-4 py-2 rounded-lg border text-sm font-bold transition ${
          todosSelecionados
            ? "bg-blue-600 border-blue-500 text-white"
            : "bg-zinc-800 border-zinc-700 text-zinc-300 hover:bg-zinc-700"
        }`}
      >
        {todosSelecionados ? "✅" : "➕"} Todos
      </button>

      {pessoas.map((nome) => {
        const ativo = selecionadas.includes(nome);

        return (
          <button
            key={nome}
            type="button"
            onClick={() => alternar(nome)}
            className={`px-4 py-2 rounded-lg border text-sm font-bold transition ${
              ativo
                ? "bg-green-600 border-green-500 text-white"
                : "bg-zinc-800 border-zinc-700 text-zinc-300 hover:bg-zinc-700"
            }`}
          >
            {ativo ? "✅" : "➕"} {nome}
          </button>
        );
      })}
    </div>
  );
}

function Card({
  titulo,
  valor,
  cor,
}: {
  titulo: string;
  valor: string;
  cor: string;
}) {
  return (
    <div className="bg-zinc-900 border border-zinc-800 p-6 rounded-xl">
      <p className="text-zinc-400">{titulo}</p>
      <p className={`text-3xl font-bold ${cor}`}>{valor}</p>
    </div>
  );
}

function TabelaLancamentos({
  titulo,
  cor,
  itens,
  novoNome,
  novoValor,
  setNovoNome,
  setNovoValor,
  adicionar,
  editar,
  excluir,
}: {
  titulo: string;
  cor: "green" | "red";
  itens: Lancamento[];
  novoNome: string;
  novoValor: string;
  setNovoNome: (valor: string) => void;
  setNovoValor: (valor: string) => void;
  adicionar: () => void;
  editar: (id: number, campo: "descricao" | "valor", valor: string) => void;
  excluir: (id: number) => void;
}) {
  const total = itens.reduce((soma, item) => soma + item.valor, 0);

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl overflow-hidden">
      <div
        className={`${cor === "green" ? "bg-green-600" : "bg-red-600"} p-3 font-bold`}
      >
        {titulo}
      </div>

      <div className="grid grid-cols-[1fr_180px_100px] bg-zinc-800 font-bold">
        <div className="p-3">Descrição</div>
        <div className="p-3">Valor</div>
        <div className="p-3">Ação</div>
      </div>

      {itens.map((item) => (
        <div
          key={item.id}
          className="grid grid-cols-[1fr_180px_100px] border-t border-zinc-800"
        >
          <input
            className="p-3 bg-zinc-900 outline-none"
            value={item.descricao}
            onChange={(e) => editar(item.id, "descricao", e.target.value)}
          />

          <input
            className="p-3 bg-zinc-900 outline-none"
            type="number"
            value={item.valor}
            onChange={(e) => editar(item.id, "valor", e.target.value)}
          />

          <button
            type="button"
            onClick={() => excluir(item.id)}
            className="text-red-400 hover:text-red-300"
          >
            Excluir
          </button>
        </div>
      ))}

      <div className="grid grid-cols-[1fr_180px_100px] border-t border-zinc-800">
        <input
          className="p-3 bg-zinc-800 outline-none"
          placeholder="Novo item"
          value={novoNome}
          onChange={(e) => setNovoNome(e.target.value)}
        />

        <input
          className="p-3 bg-zinc-800 outline-none"
          placeholder="Valor"
          value={novoValor}
          onChange={(e) => setNovoValor(e.target.value)}
        />

        <button
          type="button"
          onClick={adicionar}
          className="bg-green-600 hover:bg-green-700 font-bold"
        >
          Add
        </button>
      </div>

      <div className="grid grid-cols-2 font-bold">
        <div className="p-3 bg-zinc-800">Total</div>
        <div
          className={`${cor === "green" ? "bg-green-700" : "bg-red-700"} p-3 text-right`}
        >
          {formatar(total)}
        </div>
      </div>
    </div>
  );
}

function Tabela({
  compras,
  excluirCompra,
}: {
  compras: Compra[];
  excluirCompra: (id: number) => void;
}) {
  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl overflow-hidden">
      <table className="w-full">
        <thead className="bg-zinc-800">
          <tr>
            <th className="p-4 text-left">Compra</th>
            <th className="p-4 text-left">Cartão</th>
            <th className="p-4 text-left">Valor</th>
            <th className="p-4 text-left">Pessoa</th>
            <th className="p-4 text-left">Parcela</th>
            <th className="p-4 text-left">Data</th>
            <th className="p-4 text-left">Ação</th>
          </tr>
        </thead>

        <tbody>
          {compras.length === 0 && (
            <tr>
              <td className="p-4 text-zinc-400" colSpan={7}>
                Nenhuma despesa cadastrada ainda.
              </td>
            </tr>
          )}

          {compras.map((item) => (
            <tr key={item.id} className="border-t border-zinc-800">
              <td className="p-4">{item.descricao}</td>
              <td className="p-4">{item.cartao}</td>
              <td className="p-4 text-red-500 font-bold">
                {formatar(item.valor)}
              </td>
              <td className="p-4">{item.pessoa}</td>
              <td className="p-4">
                {item.parcelaAtual || 1}/{item.parcelas}
              </td>
              <td className="p-4">{item.data}</td>
              <td className="p-4">
                <button
                  type="button"
                  onClick={() => excluirCompra(item.id)}
                  className="text-red-400 hover:text-red-300"
                >
                  Excluir
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}