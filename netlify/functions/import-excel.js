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
    console.log('Token existe?', !!process.env.NOTION_TOKEN);
    console.log('Database ID:', process.env.NOTION_DATABASE_ID);
    
    if (!process.env.NOTION_TOKEN || !process.env.NOTION_DATABASE_ID) {
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ error: 'Variáveis de ambiente não configuradas' })
      };
    }

    const notion = new Client({ auth: process.env.NOTION_TOKEN });
    const DATABASE_ID = process.env.NOTION_DATABASE_ID;
    
    console.log('Notion client inicializado');

    // Parse arquivo
    console.log('Parseando multipart...');
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

    const rows = data.slice(1, 1001); // Máximo 1000 linhas (ignora cabeçalho)
console.log('Total de linhas (limitado a 1000):', rows.length);
    
    let imported = 0;
    let skipped = 0;
    let errors = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      
      try {
        const cnpj = row.A?.toString().trim() || '';
        const empresa = row.C?.toString().trim() || '';
        const telefone1 = row.M?.toString().trim() || '';
        const telefone2 = row.N?.toString().trim() || '';
        const email = row.O?.toString().trim() || '';

        if (!cnpj && !empresa) {
          skipped++;
          continue;
        }

        console.log(`Linha ${i + 2}: ${empresa || cnpj}`);

        const cleanPhone = (phone) => {
          if (!phone) return null;
          const cleaned = phone.replace(/\D/g, '');
          return cleaned.length > 0 ? cleaned : null;
        };

        const tel1 = cleanPhone(telefone1);
        const tel2 = cleanPhone(telefone2);

        // FORMATO CORRETO PARA NOTION
        const properties = {
          'Empresa': {
            title: [{ text: { content: empresa } }]  // CORRIGIDO: title ao invés de rich_text
          },
          'Status': {
            status: { name: 'Entrada' }  // CORRIGIDO: status ao invés de select
          }
        };

        // Adicionar CNPJ se existir
        if (cnpj) {
          properties['CNPJ'] = {
            rich_text: [{ text: { content: cnpj } }]
          };
        }

        // Adicionar telefones se existirem
        if (tel1) {
          properties['Telefone'] = { phone_number: tel1 };
        }
        if (tel2) {
          properties['Telefone 2'] = { phone_number: tel2 };
        }

        // Adicionar email se existir
        if (email) {
          properties['Email'] = { email: email };
        }

        await notion.pages.create({
          parent: { database_id: DATABASE_ID },
          properties: properties
        });

        imported++;
        console.log(`✓ Linha ${i + 2} importada`);

      } catch (error) {
        console.error(`✗ Erro na linha ${i + 2}:`, error.message);
        errors.push(`Linha ${i + 2}: ${error.message}`);
        
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
        errors: errors.length > 0 ? errors.slice(0, 5) : undefined
      })
    };

  } catch (error) {
    console.error('ERRO FATAL:', error);
    
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ 
        success: false,
        error: error.message
      })
    };
  }
};
