/* pauta-chat.js — aparência de chat (foto, nome, hora) para o histórico de orientações da pauta.
 * Usado pelo admin.html e pelo cliente/dashboard.html. Só formata; quem chama passa as funções
 * de escape e de avatar de cada tela. O texto gravado não muda: "[dd/mm/aa hh:mm · Origem · Nome] texto". */
(function (global) {
  'use strict';

  function norm(s) {
    return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  /** "20/11/26 16:01 · Escritório · Samany Cutrim" -> { hora, origem, nome }
   *  Mensagens antigas ("… · Escritório" / "… · Admin") vêm sem nome. */
  function parseMeta(meta) {
    var partes = String(meta || '').split(' · ').map(function (p) { return p.trim(); }).filter(Boolean);
    var hora = partes[0] || '';
    var origem = '', nome = '';
    if (partes[1] && /^escrit/i.test(partes[1])) { origem = 'Escritório'; nome = partes.slice(2).join(' · '); }
    else if (partes[1] && /^admin/i.test(partes[1])) { origem = 'Admin'; nome = partes.slice(2).join(' · '); }
    else if (partes.length > 1) { nome = partes.slice(1).join(' · '); }
    return { hora: hora, origem: origem, nome: nome };
  }

  /** Procura a pessoa da Equipe pelo nome (sem acento/caixa) ou sigla. */
  function acharMembro(nome, equipe) {
    var n = norm(nome);
    if (!n) return null;
    var lista = equipe || [];
    for (var i = 0; i < lista.length; i++) {
      if (norm(lista[i].nome) === n || (lista[i].sigla && norm(lista[i].sigla) === n)) return lista[i];
    }
    return null;
  }

  /** HTML de uma mensagem em formato de chat.
   *  o = { meta, euNome, equipe, avatar(c,size), esc(str), textoHtml, idTexto, acoesHtml, mono }
   *  Mensagem própria vai à direita (avatar à direita, balão azul claro); as demais à esquerda. */
  function mensagemHtml(o) {
    var esc = o.esc;
    var info = parseMeta(o.meta);
    var membro = acharMembro(info.nome, o.equipe);
    var exibe = info.nome || info.origem || 'Mensagem';
    var minha = !!(o.euNome && info.nome && norm(info.nome) === norm(o.euNome));
    var av = o.avatar(membro || { nome: exibe }, 32);
    var fonte = o.mono || 'monospace';
    var balao = 'background:' + (minha ? '#E8EEF9' : '#fff') + ';border:1px solid rgba(13,43,94,.14);'
      + 'border-radius:' + (minha ? '10px 2px 10px 10px' : '2px 10px 10px 10px') + ';padding:6px 10px;flex:1;min-width:0;max-width:calc(100% - 44px)';
    return '<div style="display:flex;gap:8px;align-items:flex-start;margin-top:8px;flex-direction:' + (minha ? 'row-reverse' : 'row') + '">'
      + '<div style="flex-shrink:0">' + av + '</div>'
      + '<div style="' + balao + '">'
      + '<div style="display:flex;align-items:baseline;gap:6px;flex-wrap:wrap;margin-bottom:3px;font-family:' + fonte + '">'
      + '<b style="font-size:11px;color:#0D2B5E">' + esc(exibe) + '</b>'
      + (info.origem && info.nome ? '<span style="font-size:8px;letter-spacing:.06em;text-transform:uppercase;color:#6B7280">' + esc(info.origem) + '</span>' : '')
      + '<span style="font-size:9px;color:#9CA3AF;margin-left:auto">' + esc(info.hora) + '</span>'
      + (o.acoesHtml || '')
      + '</div>'
      + '<div id="' + o.idTexto + '" style="font-family:' + fonte + ';font-size:11px;line-height:1.45;color:#0D2B5E;white-space:pre-wrap;overflow-wrap:anywhere">' + o.textoHtml + '</div>'
      + '</div></div>';
  }

  global.FalawChat = { parseMeta: parseMeta, acharMembro: acharMembro, mensagemHtml: mensagemHtml, _norm: norm };
})(typeof window !== 'undefined' ? window : globalThis);
