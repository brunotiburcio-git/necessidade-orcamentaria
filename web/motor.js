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

  function calcular(D, p) {
    var c = D.col, n = c.repasse.length, g = p.geral;
    var out = { ritmoSN: new Array(n), inicial: new Array(n), pos: new Array(n),
                simp: new Array(n), final: new Array(n), bv: new Array(n) };
    for (var i = 0; i < n; i++) {
      var aExec = c.aExec[i], cls = c.cls[i], disp = c.disp[i], aEmp = c.aEmp[i];
      // BF:BI e BJ: só a coluna da secretaria do contrato é diferente de zero
      var ritmoSN = 0;
      for (var s = 0; s < D.secretarias.length; s++) {
        var v = 0;
        if (aExec > 0 && c.sec[i] === s) {
          var sec = D.secretarias[s];
          v = ritmo(c, i, p.bln_secretaria[sec] ? p.secretaria[sec] : g);
        }
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
      var bs = c.bs[i], emp = c.emp[i], simpIni = c.repasse[i] * g.simplif, simp;
      if (bs === BS_NAO || bs === BS_SEM_EXEC) simp = 0;
      else if (bs === BS_SIM && aEmp === 0) simp = 0;
      else if (bs === BS_SIM && emp > simpIni) simp = 0;
      else if (bs === BS_SIM && emp < simpIni && simpIni - emp >= aEmp) simp = aEmp;
      else if (bs === BS_SIM && emp < simpIni && simpIni - emp < aEmp) simp = simpIni - emp;
      else simp = 0;
      out.ritmoSN[i] = ritmoSN; out.inicial[i] = inicial; out.pos[i] = pos; out.simp[i] = simp;
      out.final[i] = pos >= simp ? pos : simp;  // max(pos, simp) do Python
      out.bv[i] = p.bln_execucao === true && c.emExec[i];  // BV
    }
    return out;
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

  raiz.Motor = { calcular: calcular, resumoSecretarias: resumoSecretarias, resumoAcoes: resumoAcoes,
                 dePct: dePct, paraPct: paraPct };
})(typeof window !== "undefined" ? window : globalThis);
