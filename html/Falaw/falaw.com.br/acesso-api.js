/* acesso-api.js — único ponto de contato com a criação/gestão de acessos da Equipe.
 *
 * Hoje cada função chama uma função SQL (supabase/migrations/20261120_equipe_acessos.sql)
 * via sb.rpc(). Para trocar por uma Edge Function basta reescrever o miolo de
 * criarUsuario / redefinirSenha / atualizarAcesso aqui: as telas só usam este objeto.
 *
 * Todas as funções recebem o cliente Supabase (sb) já criado e devolvem
 * { ok: true, data } ou { ok: false, erro: '<codigo>', mensagem: '<texto pt-BR>' }.
 * Nenhuma senha é guardada ou registrada aqui.
 */
(function (global) {
  'use strict';

  var MENSAGENS = {
    sem_permissao: 'Você não tem permissão para esta ação.',
    email_invalido: 'E-mail inválido.',
    senha_curta: 'A senha precisa ter pelo menos 8 caracteres.',
    tipo_invalido: 'Tipo de acesso inválido.',
    email_ja_existe: 'Já existe um usuário com este e-mail.',
    pessoa_ja_tem_acesso: 'Esta pessoa já tem acesso criado.',
    acesso_inexistente: 'Esta pessoa ainda não tem acesso criado.',
    rede: 'Não foi possível falar com o servidor. Tente novamente.'
  };

  function falha(error) {
    var texto = String((error && (error.message || error.details)) || error || '');
    var codigo = Object.keys(MENSAGENS).filter(function (k) { return texto.indexOf(k) >= 0; })[0] || 'rede';
    return { ok: false, erro: codigo, mensagem: MENSAGENS[codigo] };
  }

  async function chamar(sb, fn, args) {
    try {
      var r = await sb.rpc(fn, args || {});
      if (r.error) return falha(r.error);
      return { ok: true, data: r.data };
    } catch (e) { return falha(e); }
  }

  var FalawAcesso = {
    /** Papel da sessão atual: { superadmin, ativo, tipo, abas, must_change, equipe_id } ou null. */
    meuAcesso: function (sb) { return chamar(sb, 'falaw_meu_acesso'); },

    /** Cria o login da pessoa com a senha padrão informada por quem cria. */
    criarUsuario: function (sb, o) {
      return chamar(sb, 'falaw_criar_usuario', {
        p_equipe_id: o.equipeId, p_email: o.email, p_senha: o.senha,
        p_tipo: o.tipo || 'dashboard', p_abas: o.abas || []
      });
    },

    /** Volta a senha da pessoa para a senha padrão informada; ela terá de trocá-la ao entrar. */
    redefinirSenha: function (sb, o) {
      return chamar(sb, 'falaw_redefinir_senha', { p_equipe_id: o.equipeId, p_senha: o.senha });
    },

    /** Ativa/desativa o login e define tipo (admin|dashboard) e abas liberadas. */
    atualizarAcesso: function (sb, o) {
      return chamar(sb, 'falaw_atualizar_acesso', {
        p_equipe_id: o.equipeId, p_ativo: !!o.ativo, p_tipo: o.tipo, p_abas: o.abas || []
      });
    },

    /** Lista os acessos já criados (para a página Equipe). */
    listarAcessos: async function (sb) {
      try {
        var r = await sb.from('equipe_acessos')
          .select('equipe_id,email,ativo,tipo,abas,must_change');
        if (r.error) return falha(r.error);
        return { ok: true, data: r.data || [] };
      } catch (e) { return falha(e); }
    },

    /** Primeiro login: grava a senha nova escolhida pela pessoa e tira a obrigação de trocar. */
    trocarSenhaPropria: async function (sb, novaSenha) {
      if (!novaSenha || novaSenha.length < 8) return falha('senha_curta');
      try {
        var u = await sb.auth.updateUser({ password: novaSenha });
        if (u.error) return falha(u.error);
        return await chamar(sb, 'falaw_concluir_troca_senha');
      } catch (e) { return falha(e); }
    }
  };

  global.FalawAcesso = FalawAcesso;
})(window);
