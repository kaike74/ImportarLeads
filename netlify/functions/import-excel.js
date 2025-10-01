const { Client } = require('@notionhq/client');
const XLSX = require('xlsx');
const multipart = require('parse-multipart-data');

exports.handler = async (event) => {
  console.log('=== INÍCIO DA FUNÇÃO ===');
  console.log('Método:', event.httpMethod);
  
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  try {
    // VALIDAR VARIÁVEIS DE AMBIENTE
    console.log('Token existe?', !!process.env.NOTION_TOKEN);
    console.log('Database ID:', process.env.NOTION_DATABASE_ID);
    
    if (!process.env.NOTION_TOKEN) {
      console.error('ERRO: NOTION_TOKEN não configurado!');
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ 
          error: 'Variável NOTION_TOKEN não configurada no Netlify' 
        })
      };
    }
    
    if (!process.env.NOTION_DATABASE_ID) {
      console.error('ERRO: NOTION_DATABASE_ID não configurado!');
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ 
          error: 'Variável NOTION_DATABASE_ID não configurada no Netlify' 
        })
      };
    }

    // Inicializar Notion
    const notion = new Client({ auth: process.env.NOTION_TOKEN });
    const DATABASE_ID = process.env.NOTION_DATABASE_ID;
    
    console.log('Notion client inicializado');

    // Parse do arquivo
    console.log('Parseando multipart...');
    const boundary = multipart.getBoundary(event.headers['content-type']);
    const parts = multipart.parse(Buffer.from(event.body, 'base64'), boundary);
    
    const filePart = parts.find(part => part.name === 'file');
    if (!filePart) {
      console.error('Arquivo não encontrado no upload');
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ error: 'Nenhum arquivo enviado' })
      };
    }
    
    console.log('Arquivo recebido, tamanho:', filePart.data.length, 'bytes');

    // Ler Excel
    console.log('Lendo Excel...');
    const workbook = XLSX.read(filePart.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    const data = XLSX.utils.sheet_to_json(sheet, { 
      header: 'A',
      defval: ''
    });

    const rows = data.slice(1); // Pula cabeçalho
    console.log('Total de linhas (sem cabeçalho):', rows.length);
    
    let imported = 0;
    let skipped = 0;
    let errors = [];

    // Processar linha por linha
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      
      try {
        const cnpj = row.A?.toString().trim() || '';
        const empresa = row.C?.toString().trim() || '';
        const telefone1 = row.M?.toString().trim() || '';
        const telefone2 = row.N?.toString().trim() || '';
        const email = row.O?.toString().trim() || '';

        // Pula linhas vazias
        if (!cnpj && !empresa) {
          skipped++;
          continue;
        }

        console.log(`Linha ${i + 2}: ${empresa || cnpj}`);

        // Limpar telefones
        const cleanPhone = (phone) => {
          if (!phone) return null;
          const cleaned = phone.replace(/\D/g, '');
          return cleaned.length > 0 ? cleaned : null;
        };

        const tel1 = cleanPhone(telefone1);
        const tel2 = cleanPhone(telefone2);

        // Criar no Notion
        await notion.pages.create({
          parent: { database_id: DATABASE_ID },
          properties: {
            'CNPJ': {
              rich_text: [{ text: { content: cnpj } }]
            },
            'Empresa': {
              rich_text: [{ text: { content: empresa } }]
            },
            'Telefone': tel1 ? {
              phone_number: tel1
            } : { phone_number: null },
            'Telefone 2': tel2 ? {
              phone_number: tel2
            } : { phone_number: null },
            'Email': email ? {
              email: email
            } : { email: null },
            'Status': {
              select: { name: 'Entrada' }
            }
          }
        });

        imported++;
        console.log(`✓ Linha ${i + 2} importada`);

      } catch (error) {
        console.error(`✗ Erro na linha ${i + 2}:`, error.message);
        errors.push(`Linha ${i + 2}: ${error.message}`);
        
        // Continua mesmo com erro
        if (errors.length > 10) {
          console.log('Muitos erros, parando...');
          break;
        }
      }
    }

    console.log('=== RESUMO ===');
    console.log('Importados:', imported);
    console.log('Pulados:', skipped);
    console.log('Erros:', errors.length);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        success: true,
        imported,
        skipped,
        errors: errors.slice(0, 5) // Primeiros 5 erros
      })
    };

  } catch (error) {
    console.error('ERRO FATAL:', error);
    console.error('Stack:', error.stack);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        success: false,
        error: error.message,
        details: error.stack
      })
    };
  }
};
