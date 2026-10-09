// Motor de cálculo da versão web: mesmas fórmulas de necorc/modelo.py (que reproduz o Excel).
// Recalcula só as colunas que dependem dos parâmetros; o resto vem pronto em D.col (necorc/artifact.py).
// A ordem das contas e das somas segue o Python/Excel para o resultado ser idêntico, bit a bit.
(function (raiz) {
  "use strict";
  var EXEC_100 = 0, DESEMB_100 = 1, EMPENH_100 = 2, SEM_EXEC = 3, COM_EXEC = 4;
  var BS_NAO = 0, BS_SEM_EXEC = 1, BS_SIM = 2;

  // _ritmo: núcleo das colunas BB e BF:BI depois do teste "a executar total > 0"
  function ritmo(c, i, g) {
    var repasse = c.repasse[i], mm = c.mm[i];
    if (!c.execSim[i]) return repasse * g.piso * g.meses;
    if (mm > g.teto) return repasse * g.meses * g.teto;
    if (mm < g.piso) return repasse * g.meses * g.piso;
    return c.ref[i] * g.meses * (1 + g.margem_exec);
  }

  // ARRED do Excel como no Python (Decimal(repr(x)).quantize, meio para longe do zero): arredonda o texto
  function arred(x, casas) {
    var m = String(x).match(/^(-?)(\d+)(?:\.(\d+))?(?:e([+-]\d+))?$/);
    if (!m) return x;
    var dig = m[2] + (m[3] || ""), ponto = m[2].length + (m[4] ? parseInt(m[4], 10) : 0);
    var manter = ponto + casas;  // algarismos que ficam
    if (manter >= dig.length) return x;
    if (manter < 0) return Number(m[1] + "0");
    var cab = dig.slice(0, manter) || "0";
    if (dig.charCodeAt(manter) - 48 >= 5) {  // soma 1 no último algarismo mantido
      var a = cab.split(""), k = a.length - 1;
      while (k >= 0 && a[k] === "9") { a[k] = "0"; k--; }
      if (k < 0) a.unshift("1"); else a[k] = String.fromCharCode(a[k].charCodeAt(0) + 1);
      cab = a.join("");
    }
    return Number(m[1] + cab + "e" + (ponto - manter));
  }

  function situacao(c, i, g) {  // BC Coluna1
    if (!(c.aExec[i] > 0)) return "situacao0";
    if (!c.execSim[i]) return "situacao1";
    if (c.mm[i] > g.teto) return "situação2";
    if (c.mm[i] < g.piso) return "situação3";
    return "situação4";
  }

  function criterio(cls, aExec, ritmoG, disp, aEmp, pctG, g) {  // BM
    if (cls === EXEC_100) return "R$0, pois já está 100% executado";
    if (cls === DESEMB_100) return "R$0, pois já está 100% desembolsado";
    if (cls === EMPENH_100) return "R$0, pois já está 100% empenhado";
    if (cls === SEM_EXEC && aExec > 0) return "Ritmo definido para sem execução menos valor disponível";
    if (cls === COM_EXEC && aExec > 0 && ritmoG - disp >= aEmp) return " Valor limitado pelo a empenhar do contrato";
    if (cls === COM_EXEC && aExec > 0 && pctG < g.piso) return "Regra do piso mínimo menos vlr disponível";
    if (cls === COM_EXEC && aExec > 0 && pctG > g.teto) return "Regra do teto máximo menos vlr disponível";
    return "Ritmo de execução estimada menos vlr disponível";
  }

  // completo = true também calcula as colunas que só aparecem no detalhe exportado
  function calcular(D, p, completo) {
    var c = D.col, n = c.repasse.length, g = p.geral, S = D.secretarias.length;
    var out = { ritmoSN: new Array(n), inicial: new Array(n), pos: new Array(n),
                simp: new Array(n), final: new Array(n), bv: new Array(n) };
    if (completo) {
      ["ritmoG", "col1", "pctG", "be", "pctSN", "bl", "crit", "simpIni"].forEach(function (k) { out[k] = new Array(n); });
      out.porSec = D.secretarias.map(function () { return new Array(n); });
    }
    for (var i = 0; i < n; i++) {
      var aExec = c.aExec[i], cls = c.cls[i], disp = c.disp[i], aEmp = c.aEmp[i], repasse = c.repasse[i];
      // BF:BI e BJ: só a coluna da secretaria do contrato é diferente de zero
      var ritmoSN = 0;
      for (var s = 0; s < S; s++) {
        var v = 0;
        if (aExec > 0 && c.sec[i] === s) {
          var sec = D.secretarias[s];
          v = ritmo(c, i, p.bln_secretaria[sec] ? p.secretaria[sec] : g);
        }
        if (completo) out.porSec[s][i] = v;
        ritmoSN += v;
      }
      // BL
      var bl;
      if (cls === EXEC_100 || cls === DESEMB_100 || cls === EMPENH_100) bl = 0;
      else if (cls === COM_EXEC && aExec > 0 && ritmoSN - disp >= aEmp) bl = aEmp;
      else bl = ritmoSN - disp;
      // BN
      var inicial = bl >= 0 ? bl : 0;
      // BQ
      var necfin = c.necfin[i], dispPos = c.dispPos[i], pos;
      if (necfin === 0) pos = inicial;
      else if (necfin > 0 && inicial >= dispPos) pos = inicial - dispPos;
      else if (necfin > 0 && inicial < dispPos) pos = 0;
      else if (necfin > 0 && dispPos < 0) pos = inicial + Math.abs(dispPos);
      else pos = 0;
      // BS e BT
      var bs = c.bs[i], emp = c.emp[i], simpIni = repasse * g.simplif, simp;
      if (bs === BS_NAO || bs === BS_SEM_EXEC) simp = 0;
      else if (bs === BS_SIM && aEmp === 0) simp = 0;
      else if (bs === BS_SIM && emp > simpIni) simp = 0;
      else if (bs === BS_SIM && emp < simpIni && simpIni - emp >= aEmp) simp = aEmp;
      else if (bs === BS_SIM && emp < simpIni && simpIni - emp < aEmp) simp = simpIni - emp;
      else simp = 0;
      out.ritmoSN[i] = ritmoSN; out.inicial[i] = inicial; out.pos[i] = pos; out.simp[i] = simp;
      out.final[i] = pos >= simp ? pos : simp;  // max(pos, simp) do Python
      out.bv[i] = p.bln_execucao === true && c.emExec[i];  // BV
      if (!completo) continue;
      // BB, BC, BD, BE (regra geral), BK, BM
      var ritmoG = aExec > 0 ? ritmo(c, i, g) : 0;
      var pctG = repasse === 0 ? 0 : arred((ritmoG / repasse) / g.meses, 4);
      var be;
      if (cls === EXEC_100 || cls === DESEMB_100 || cls === EMPENH_100) be = 0;
      else if (cls === SEM_EXEC && aExec > 0) be = ritmoG - disp;
      else if (cls === COM_EXEC && aExec > 0 && ritmoG - disp >= aEmp) be = aEmp;
      else if (cls === COM_EXEC && aExec > 0 && pctG < g.piso) be = (c.ref[i] * g.piso * g.meses) - disp;
      else if (cls === COM_EXEC && aExec > 0 && pctG > g.teto) be = (c.ref[i] * g.teto * g.meses) - disp;
      else be = ritmoG - disp;
      out.ritmoG[i] = ritmoG; out.col1[i] = situacao(c, i, g); out.pctG[i] = pctG; out.be[i] = be;
      out.pctSN[i] = repasse === 0 ? 0 : arred((ritmoSN / repasse) / g.meses, 4);
      out.bl[i] = bl; out.crit[i] = criterio(cls, aExec, ritmoG, disp, aEmp, pctG, g); out.simpIni[i] = simpIni;
    }
    return out;
  }

  // Aba "detalhe contratos": todas as colunas da view, na ordem do Python (pd.DataFrame(linhas))
  function detalhe(D, p) {
    var r = calcular(D, p, true), d = D.detalhe, recalc = {
      "Ritmo de execução parâmetros": r.ritmoG, "Coluna1": r.col1, "% exec mensal parâmetros": r.pctG,
      "necessidade orçamentária preliminar": r.be, "ritmo parametros SNs": r.ritmoSN,
      "% exec mensal parâmetros SNs": r.pctSN, "necessidade preliminar regra secretaria": r.bl,
      "critério inicial aplicado": r.crit, "necessidade orçamento inicial": r.inicial,
      "necessidade orçamento pós NecFin": r.pos, "Simplificado inicial": r.simpIni,
      "simplificado final": r.simp, "necessidade orçamento final": r.final, "bln_execução (situação obra)": r.bv
    };
    D.secretarias.forEach(function (s, k) { recalc["ritmo parametros " + s] = r.porSec[k]; });
    var colunas = d.colunas.map(function (nome) {
      var v = recalc[nome] || d.fixas[nome];
      if (!v) throw new Error("coluna sem dados: " + nome);
      return v;
    });
    return { colunas: d.colunas, valores: colunas, datas: d.datas, r: r };
  }

  function somaLinhas(valores, linhas, filtro) {  // SOMASES: soma na ordem das linhas
    var t = 0;
    for (var k = 0; k < linhas.length; k++) {
      var i = linhas[k];
      if (!filtro || filtro(i)) t += valores[i];
    }
    return t;
  }

  function somaLista(valores) { var t = 0; for (var k = 0; k < valores.length; k++) t += valores[k]; return t; }

  function resumoSecretarias(D, r, p) {  // K11:P18 e K20:P20 (milhões)
    var c = D.col;
    var itens = [
      ["Ritmo de execução", r.ritmoSN], ["Valor disponível", c.disp],
      ["Necessidade orçamentária preliminar", r.inicial], ["Necessidade Financeira CAIXA", c.necfin],
      ["Necessidade orçamentária após CAIXA", r.pos], ["Necessidade orçamentária simplificado", r.simp],
      ["Necessidade orçamentária final", r.final]
    ];
    var out = itens.map(function (it) {
      var vals = D.linhas_resumo.map(function (ls) { return somaLinhas(it[1], ls) / 1000000; });
      return { rotulo: it[0], valores: vals, total: somaLista(vals) };
    });
    var vals20 = D.linhas_resumo.map(function (ls) {
      return somaLinhas(r.final, ls, function (i) { return r.bv[i] === p.bln_execucao; }) / 1000000;
    });
    out.push({ rotulo: "Nec orçamentária final (bln_execucao)", valores: vals20, total: somaLista(vals20) });
    return out;
  }

  function resumoAcoes(D, r) {  // K25:P47 (milhões)
    var itens = D.acoes.map(function (a) {
      var nec = somaLinhas(r.final, a.linhas) / 1000000, disp = a.disponivel_loa;
      return { descricao: a.descricao, codigo: a.codigo, nec: nec, disp: disp,
               dentro: nec > disp ? "Não" : "Sim", saldo: disp - nec };
    });
    return {
      itens: itens,
      totalNec: somaLista(itens.map(function (i) { return i.nec; })),
      totalDisp: somaLista(itens.map(function (i) { return i.disp; })),
      faltando: somaLista(itens.filter(function (i) { return i.dentro === "Não"; }).map(function (i) { return i.saldo; })),
      sobrando: somaLista(itens.filter(function (i) { return i.dentro === "Sim"; }).map(function (i) { return i.saldo; }))
    };
  }

  // 0,6 (%) -> 0,006 exatamente como o Python (Decimal(repr(x)) / 100): desloca a vírgula no texto
  function dePct(pct) {
    var s = String(pct), m = s.match(/^(-?)(\d+)(?:\.(\d+))?(?:e([+-]\d+))?$/);
    if (!m) return pct / 100;
    var exp = (m[4] ? parseInt(m[4], 10) : 0) - 2;
    return Number(m[1] + m[2] + (m[3] ? "." + m[3] : "") + "e" + exp);
  }
  function paraPct(fracao) {
    var s = String(fracao), m = s.match(/^(-?)(\d+)(?:\.(\d+))?(?:e([+-]\d+))?$/);
    if (!m) return fracao * 100;
    var exp = (m[4] ? parseInt(m[4], 10) : 0) + 2;
    return Number(m[1] + m[2] + (m[3] ? "." + m[3] : "") + "e" + exp);
  }

  raiz.Motor = { calcular: calcular, detalhe: detalhe, arred: arred, resumoSecretarias: resumoSecretarias, resumoAcoes: resumoAcoes,
                 dePct: dePct, paraPct: paraPct };
})(typeof window !== "undefined" ? window : globalThis);
