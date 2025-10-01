const { Client } = require('@notionhq/client');
const XLSX = require('xlsx');
const multipart = require('parse-multipart-data');

const notion = new Client({ auth: process.env.NOTION_TOKEN });
const DATABASE_ID = process.env.NOTION_DATABASE_ID;

exports.handler = async (event) => {
  // CORS headers
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };

  // Handle OPTIONS
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  try {
    // Parse multipart data
    const boundary = multipart.getBoundary(event.headers['content-type']);
    const parts = multipart.parse(Buffer.from(event.body, 'base64'), boundary);
    
    const filePart = parts.find(part => part.name === 'file');
    if (!filePart) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ error: 'Nenhum arquivo enviado' })
      };
    }

    // Ler Excel
    const workbook = XLSX.read(filePart.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    // Converter para JSON
    const data = XLSX.utils.sheet_to_json(sheet, { 
      header: 'A',
      defval: ''
    });

    // Pular cabeçalho (linha 1)
    const rows = data.slice(1);
    
    let imported = 0;
    let errors = [];

    // Processar cada linha
    for (const row of rows) {
      try {
        const cnpj = row.A || '';
        const empresa = row.C || '';
        const telefone1 = row.M || '';
        const telefone2 = row.N || '';
        const email = row.O || '';

        // Validação básica
        if (!cnpj && !empresa) continue; // Pula linhas vazias

        // Limpar telefones
        const cleanPhone = (phone) => {
          return phone.toString().replace(/\D/g, '');
        };

        // Criar no Notion
        await notion.pages.create({
          parent: { database_id: DATABASE_ID },
          properties: {
            'CNPJ': {
              rich_text: [{ text: { content: cnpj.toString() } }]
            },
            'Empresa': {
              rich_text: [{ text: { content: empresa.toString() } }]
            },
            'Telefone': {
              phone_number: cleanPhone(telefone1) || null
            },
            'Telefone 2': {
              phone_number: cleanPhone(telefone2) || null
            },
            'Email': {
              email: email || null
            },
            'Status': {
              select: { name: 'Entrada' }
            }
          }
        });

        imported++;
      } catch (error) {
        errors.push(`Erro na linha: ${error.message}`);
      }
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        imported,
        errors: errors.length > 0 ? errors : undefined
      })
    };

  } catch (error) {
    console.error('Erro:', error);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: error.message })
    };
  }
};
