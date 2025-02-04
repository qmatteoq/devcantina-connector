import { GraphError } from '@microsoft/microsoft-graph-client';
import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import { config } from './config.js';
import { client } from './graphClient.js';
import { parseStringPromise } from 'xml2js';

// Represents the document to import
interface Document {
  // Document title
  title: string;
  // Document content. Can be plain-text or HTML
  content: string;
  // URL to the document in the external system
  url: string;
  // URL to the document icon. Required by Microsoft Copilot for Microsoft 365
  iconUrl: string;
  //Document description
  description: string;
  // URL to the post image
  image: string;
}

function stripHtmlTags(str: string): string {
  return str.replace(/<\/?[^>]+(>|$)/g, "");
}

async function extract(): Promise<Document[]> {
  const rssUrl = 'https://www.developerscantina.com/index.xml';
  const response = await fetch(rssUrl);
  const rssText = await response.text();
  const rssJson = await parseStringPromise(rssText);

  const documents: Document[] = rssJson.rss.channel[0].item.map((item: any) => {
    const description = item.description[0];
    let imageUrl = '';

    // Extract image URL from description
    const imgTagMatch = description.match(/<img[^>]+src="([^">]+)"/);
    if (imgTagMatch && imgTagMatch[1]) {
      imageUrl = imgTagMatch[1];
    }

    const cleanDescription = stripHtmlTags(description);

    return {
      title: item.title[0],
      content: description,
      url: item.link[0],
      iconUrl: item['media:thumbnail'] ? item['media:thumbnail'][0].$.url : '',
      description: cleanDescription.length >= 250 ? "..." + cleanDescription.substring(250, 600) + "..." : " ..." + cleanDescription + "...",
      image: imageUrl
    };
  });

  return documents;
}

function getDocId(doc: Document): string {
  try {
    const url = new URL(doc.url);
    const pathSegments = url.pathname.split('/');
    console.log('pathSegments:', pathSegments);
    console.log('pathSegments length:', pathSegments.length);
    const slug = pathSegments[pathSegments.length - 2];
    console.log('slug:', slug);
    return slug;
  } catch (error) {
    console.error('Invalid URL:', doc.url);
    return '';
  }
}

function transform(documents: Document[]): ExternalConnectors.ExternalItem[] {
  return documents.map(doc => {
    const docId = getDocId(doc);
    return {
      id: docId,
      properties: {
        // Add properties as defined in the schema in config.ts
        title: doc.title ?? '',
        url: doc.url,
        iconUrl: doc.iconUrl,
        description: doc.description,
        imageUrl: doc.image
      },
      content: {
        value: doc.content ?? '',
        type: 'text'
      },
      acl: [
        {
          accessType: 'grant',
          type: 'everyone',
          value: 'everyone'
        }
      ]
    } as ExternalConnectors.ExternalItem
  });
}

async function load(externalItems: ExternalConnectors.ExternalItem[]) {
  const { id } = config.connection;
  for (const doc of externalItems) {
    try {
      console.log(`Loading ${doc.id}...`);
      await client
        .api(`/external/connections/${id}/items/${doc.id}`)
        .header('content-type', 'application/json')
        .put(doc);
      console.log('  DONE');
    }
    catch (e) {
      const graphError = e as GraphError;
      console.error(`Failed to load ${doc.id}: ${graphError.message}`);
      if (graphError.body) {
        console.error(`${JSON.parse(graphError.body)?.innerError?.message}`);
      }
      return;
    }
  }
}

export async function loadContent() {
  const content = await extract();
  const transformed = transform(content);
  await load(transformed);
}

loadContent();