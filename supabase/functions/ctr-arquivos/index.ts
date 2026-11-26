// ════════════════════════════════════════════════════════════════════════
// Gestao de Contratos — ponte entre o sistema e o SharePoint (Supabase Edge Function)
//
// Nome da funcao no Supabase: ctr-arquivos
// Segredos (Supabase > Edge Functions > Secrets), nunca no codigo:
//   MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET  -> aplicativo "Sistema de Gestao de Contratos" no Entra
//   SP_SITE_ID                                    -> id do site GestaoContratos no SharePoint
//   SP_PASTA_RAIZ (opcional, padrao "Contratos")  -> pasta raiz dentro da biblioteca Documentos
// SUPABASE_URL, SUPABASE_ANON_KEY e SUPABASE_SERVICE_ROLE_KEY o Supabase ja fornece.
//
// O navegador nunca recebe a senha da Microsoft. Para cada pedido a funcao:
//   1) descobre quem e o usuario pelo login dele;
//   2) pergunta ao banco o nivel dele naquele contrato (mesma regra das telas);
//   3) so entao fala com o SharePoint.
// Acoes: iniciar (envio), registrar (envio terminou), ver, baixar, testar (so admin).
// ════════════════════════════════════════════════════════════════════════
import { createClient } from "npm:@supabase/supabase-js@2.45.4";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const GRAPH = "https://graph.microsoft.com/v1.0";
const MAX_BYTES = 250 * 1024 * 1024;
// tipos de documento (iguais aos das telas); os com validade exigem a data
const CATEGORIAS: Record<string, string> = {
  "contrato-assinado": "Contrato principal assinado",
  "aditivo": "Aditivo",
  "minuta": "Minuta",
  "anexo": "Anexo / planilha",
  "renovacao": "Renovação",
  "distrato": "Distrato / rescisão",
  "correspondencia": "Correspondência / e-mail",
  "garantia": "Garantia (caução, fiança, seguro-fiança)",
  "cnd-federal": "Certidão negativa federal (Receita e PGFN)",
  "crf-fgts": "Certificado de regularidade do FGTS",
  "cndt": "Certidão negativa de débitos trabalhistas",
  "autorizacao-pf": "Autorização da Polícia Federal (vigilância)",
  "avcb": "AVCB / certificado do Corpo de Bombeiros",
  "alvara": "Alvará de funcionamento",
  "seguro": "Apólice de seguro",
  "licenca": "Licença sanitária ou ambiental",
  "outro-validade": "Outro documento com validade",
};
const COM_VALIDADE = new Set(["garantia", "cnd-federal", "crf-fgts", "cndt", "autorizacao-pf", "avcb", "alvara", "seguro", "licenca", "outro-validade"]);
const dataValida = (s: unknown) => typeof s === "string" && /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(Date.parse(s + "T00:00:00Z"));

class Falha extends Error { constructor(public status: number, msg: string) { super(msg); } }
const env = (k: string, padrao?: string) => {
  const v = Deno.env.get(k) ?? padrao;
  if (!v) throw new Falha(500, `Configuração incompleta no servidor: falta o segredo ${k}.`);
  return v;
};
const resposta = (corpo: unknown, status = 200) =>
  new Response(JSON.stringify(corpo), { status, headers: { ...CORS, "Content-Type": "application/json" } });

// ── nomes de pasta e arquivo aceitos pelo SharePoint ──
const limpar = (s: string, max = 80) =>
  String(s ?? "").normalize("NFC").replace(/["*:<>?/\\|#%~&{}]+/g, " ").replace(/\s+/g, " ")
    .trim().replace(/^[.\s]+|[.\s]+$/g, "").slice(0, max).trim() || "sem-nome";
const slug = (s: string, max = 40) =>
  String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Za-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "").slice(0, max).replace(/-+$/, "") || "sem-nome";
const extensao = (nome: string) => {
  const m = /\.([A-Za-z0-9]{1,8})$/.exec(nome || "");
  return m ? m[1].toLowerCase() : "bin";
};
const codificar = (caminho: string) => caminho.split("/").map(encodeURIComponent).join("/");

// ── Microsoft Graph (token do aplicativo, guardado enquanto vale) ──
let token: { valor: string; expira: number } | null = null;
async function tokenGraph(): Promise<string> {
  if (token && token.expira > Date.now() + 60_000) return token.valor;
  const r = await fetch(`https://login.microsoftonline.com/${env("MS_TENANT_ID")}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: env("MS_CLIENT_ID"),
      client_secret: env("MS_CLIENT_SECRET"),
      scope: "https://graph.microsoft.com/.default",
    }),
  });
  const j = await r.json();
  if (!r.ok) {
    console.error("token microsoft", r.status, j?.error, j?.error_codes);
    throw new Falha(502, "O SharePoint recusou a senha do sistema (segredo vencido ou errado). Avise o administrador.");
  }
  token = { valor: j.access_token, expira: Date.now() + (j.expires_in ?? 3000) * 1000 };
  return token.valor;
}
async function graph(caminho: string, init: RequestInit = {}, aceitar404 = false) {
  const r = await fetch(GRAPH + caminho, {
    ...init,
    headers: { Authorization: `Bearer ${await tokenGraph()}`, "Content-Type": "application/json", ...(init.headers ?? {}) },
  });
  if (aceitar404 && r.status === 404) return null;
  const texto = await r.text();
  const j = texto ? JSON.parse(texto) : {};
  if (!r.ok) {
    console.error("graph", init.method ?? "GET", caminho.split("?")[0].slice(0, 120), r.status, j?.error?.code);
    if (r.status === 403) throw new Falha(502, "O sistema não tem permissão no site do SharePoint. Avise o administrador.");
    throw new Falha(502, `O SharePoint respondeu com erro (${r.status}). Tente de novo em alguns minutos.`);
  }
  return j;
}
const site = () => `/sites/${env("SP_SITE_ID")}/drive`;

// cria cada pasta do caminho que ainda nao existe
async function garantirPasta(caminho: string) {
  const partes = caminho.split("/");
  let atual = "";
  for (const parte of partes) {
    const pai = atual;
    atual = atual ? `${atual}/${parte}` : parte;
    const existe = await graph(`${site()}/root:/${codificar(atual)}?$select=id,folder`, {}, true);
    if (existe) continue;
    const destino = pai ? `${site()}/root:/${codificar(pai)}:/children` : `${site()}/root/children`;
    try {
      await graph(destino, {
        method: "POST",
        body: JSON.stringify({ name: parte, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
      });
    } catch (e) {
      // outra pessoa pode ter criado a mesma pasta no mesmo segundo
      if (!(await graph(`${site()}/root:/${codificar(atual)}?$select=id`, {}, true))) throw e;
    }
  }
}

// ── Supabase ──
const admin = () => createClient(env("SUPABASE_URL"), env("SUPABASE_SERVICE_ROLE_KEY"), { auth: { persistSession: false } });

async function nivel(usuario: ReturnType<typeof createClient>, contrato: number): Promise<number> {
  const { data, error } = await usuario.rpc("ctr_nivel_contrato", { p_id: contrato });
  if (error) { console.error("nivel", error.message); throw new Falha(500, "Não foi possível conferir sua permissão."); }
  return Number(data) || 0;
}
async function pastaDoContrato(db: ReturnType<typeof createClient>, contrato: number) {
  const { data: c, error } = await db.from("ctr_contratos")
    .select("id,numero,empresa,categoria,natureza_cliente,fornecedor_nome").eq("id", contrato).single();
  if (error || !c) throw new Falha(404, "Contrato não encontrado.");
  const { data: emp } = await db.from("empresas").select("nome").eq("slug", c.empresa).single();
  const { data: cat } = await db.from("ctr_categorias").select("nome").eq("id", c.categoria).maybeSingle();
  // empresa que cadastra pela natureza do ERP: a pasta e a natureza (o padrao dela); senao, a categoria
  const { data: nat } = c.natureza_cliente == null ? { data: null } :
    await db.from("ctr_naturezas").select("descricao").eq("empresa", c.empresa).eq("codigo", c.natureza_cliente).maybeSingle();
  const nomeCategoria = nat ? `${c.natureza_cliente} - ${String(nat.descricao).replace(/^#\s*/, "")}` : (cat?.nome ?? "Sem categoria");
  const pasta = [
    limpar(env("SP_PASTA_RAIZ", "Contratos"), 40),
    limpar(emp?.nome ?? c.empresa, 50),
    limpar(nomeCategoria, 70),
    limpar(`${c.numero} - ${c.fornecedor_nome}`, 90),
  ].join("/");
  return { c: { ...c, nomeCategoria }, pasta };
}
async function historico(db: ReturnType<typeof createClient>, contrato: number, empresa: string, usuario: string, acao: string, detalhe: string) {
  const { error } = await db.from("ctr_historico").insert({ contrato, empresa, usuario, acao, detalhe });
  if (error) console.error("historico", error.message);
}
async function documento(db: ReturnType<typeof createClient>, id: number) {
  const { data, error } = await db.from("ctr_documentos").select("id,contrato,nome,sp_item").eq("id", id).single();
  if (error || !data) throw new Falha(404, "Arquivo não encontrado.");
  return data;
}

// ════════ Atendimento ════════
Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    if (req.method !== "POST") throw new Falha(405, "Use POST.");
    const auth = req.headers.get("Authorization") ?? "";
    if (!auth.startsWith("Bearer ")) throw new Falha(401, "Entre no sistema de novo.");
    const usuario = createClient(env("SUPABASE_URL"), env("SUPABASE_ANON_KEY"), {
      global: { headers: { Authorization: auth } }, auth: { persistSession: false },
    });
    const { data: u, error: eu } = await usuario.auth.getUser(auth.slice(7));
    if (eu || !u?.user) throw new Falha(401, "Sua sessão expirou. Entre no sistema de novo.");
    const uid = u.user.id;
    const db = admin();
    const p = await req.json().catch(() => ({}));
    const acao = String(p.acao ?? "");

    if (acao === "testar") {
      const { data: ehAdmin } = await usuario.rpc("eh_admin");
      if (!ehAdmin) throw new Falha(403, "Só o administrador pode testar a conexão.");
      const raiz = limpar(env("SP_PASTA_RAIZ", "Contratos"), 40);
      await garantirPasta(raiz);
      const d = await graph(`${site()}?$select=webUrl,quota`);
      return resposta({ ok: true, pasta: raiz, webUrl: d.webUrl, usado: d.quota?.used, total: d.quota?.total });
    }

    if (acao === "iniciar") {
      const contrato = Number(p.contrato);
      const categoria = String(p.categoria ?? "");
      const nomeOriginal = String(p.nome_original ?? "").slice(0, 200);
      const tamanho = Number(p.tamanho) || 0;
      if (!CATEGORIAS[categoria]) throw new Falha(400, "Escolha o que é o arquivo.");
      if (tamanho <= 0) throw new Falha(400, "O arquivo está vazio.");
      if (tamanho > MAX_BYTES) throw new Falha(400, "O arquivo passa de 250 MB.");
      if (await nivel(usuario, contrato) < 3) throw new Falha(403, "Você não tem permissão para enviar arquivos neste contrato.");
      const { c, pasta } = await pastaDoContrato(db, contrato);
      const { count } = await db.from("ctr_documentos").select("id", { count: "exact", head: true })
        .eq("contrato", contrato).eq("categoria", categoria);
      const versao = (count ?? 0) + 1;
      const nome = `${c.numero}_${slug(c.nomeCategoria, 30)}_${slug(c.fornecedor_nome)}_${categoria}_v${versao}.${extensao(nomeOriginal)}`;
      await garantirPasta(pasta);
      const sessao = await graph(`${site()}/root:/${codificar(`${pasta}/${nome}`)}:/createUploadSession`, {
        method: "POST",
        body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
      });
      return resposta({ uploadUrl: sessao.uploadUrl, nome, pasta, versao });
    }

    if (acao === "registrar") {
      const contrato = Number(p.contrato);
      const categoria = String(p.categoria ?? "");
      if (!CATEGORIAS[categoria]) throw new Falha(400, "Categoria inválida.");
      const validade = p.validade ? String(p.validade) : null;
      if (COM_VALIDADE.has(categoria) && !dataValida(validade)) throw new Falha(400, "Informe a data de validade do documento.");
      if (validade && !dataValida(validade)) throw new Falha(400, "Data de validade inválida.");
      if (await nivel(usuario, contrato) < 3) throw new Falha(403, "Você não tem permissão para enviar arquivos neste contrato.");
      const { c, pasta } = await pastaDoContrato(db, contrato);
      const item = await graph(`${site()}/items/${encodeURIComponent(String(p.item_id ?? ""))}?$select=id,name,size,file,parentReference`, {}, true);
      const caminhoPai = decodeURIComponent(String(item?.parentReference?.path ?? "")).replace(/^.*?root:\/?/, "");
      if (!item?.file || caminhoPai !== pasta || !String(item.name).startsWith(`${c.numero}_`)) {
        throw new Falha(400, "O arquivo enviado não foi encontrado na pasta do contrato.");
      }
      const versao = Number((/_v(\d+)[^._]*\.[^.]+$/.exec(item.name) ?? [])[1]) || 1;
      const { data: doc, error } = await db.from("ctr_documentos").insert({
        contrato, categoria, validade: COM_VALIDADE.has(categoria) ? validade : null, nome: item.name, nome_original: String(p.nome_original ?? "").slice(0, 200),
        tamanho: item.size, tipo_mime: item.file?.mimeType ?? null, versao, sp_item: item.id,
        sp_caminho: `${pasta}/${item.name}`, enviado_por: uid,
      }).select("id").single();
      if (error) {
        if (/duplicate|unique/i.test(error.message)) throw new Falha(409, "Este arquivo já foi registrado.");
        console.error("registrar", error.message); throw new Falha(500, "O arquivo subiu, mas não foi possível registrar no sistema.");
      }
      await historico(db, contrato, c.empresa, uid, "arquivo enviado", `${CATEGORIAS[categoria]}${COM_VALIDADE.has(categoria) ? ` (válido até ${validade!.split("-").reverse().join("/")})` : ""}: ${item.name}`);
      return resposta({ ok: true, id: doc.id, nome: item.name, caminho: `${pasta}/${item.name}` });
    }

    if (acao === "ver" || acao === "baixar") {
      const doc = await documento(db, Number(p.documento));
      const n = await nivel(usuario, doc.contrato);
      if (n < (acao === "baixar" ? 2 : 1)) throw new Falha(403, acao === "baixar" ? "Seu acesso permite só visualizar." : "Você não tem acesso a este contrato.");
      const { data: c } = await db.from("ctr_contratos").select("empresa").eq("id", doc.contrato).single();
      let url: string;
      if (acao === "baixar") {
        const item = await graph(`${site()}/items/${encodeURIComponent(doc.sp_item)}`);
        url = item["@microsoft.graph.downloadUrl"];
      } else {
        const prev = await graph(`${site()}/items/${encodeURIComponent(doc.sp_item)}/preview`, { method: "POST", body: "{}" });
        url = prev.getUrl;
      }
      if (!url) throw new Falha(502, "O SharePoint não devolveu o link do arquivo.");
      await historico(db, doc.contrato, c?.empresa ?? "", uid, acao === "baixar" ? "arquivo baixado" : "arquivo visualizado", doc.nome);
      return resposta({ url, nome: doc.nome });
    }

    throw new Falha(400, "Ação desconhecida.");
  } catch (e) {
    if (e instanceof Falha) return resposta({ erro: e.message }, e.status);
    console.error("erro inesperado", (e as Error)?.message);
    return resposta({ erro: "Erro inesperado no servidor. Tente de novo." }, 500);
  }
});
