import { ExternalConnectors } from '@microsoft/microsoft-graph-types';

export const config = {
  connection: {
    id: 'devcantina',
    name: 'The Developers Cantina Connector',
    description: 'Imports data from the Developers Cantina blog',
    activitySettings: {
      // URL to item resolves track activity such as sharing external items.
      // The recorded activity is used to improve search relevance.
      urlToItemResolvers: [
        {
          urlMatchInfo: {
            baseUrls: [
              'https://www.developerscantina.com'
            ],
            urlPattern: '/p/(?<slug>[^/]+)'
          },
          itemId: '{slug}',
          priority: 1
        } as ExternalConnectors.ItemIdResolver
      ]
    },
    searchSettings: {
      searchResultTemplates: [
        {
          id: 'devcantina',
          priority: 1,
          layout: {}
        }
      ]
    },
    // https://learn.microsoft.com/graph/connecting-external-content-manage-schema
    schema: {
      baseType: 'microsoft.graph.externalItem',
      // Add properties as needed
      properties: [
        {
          name: 'title',
          type: 'string',
          isQueryable: true,
          isSearchable: true,
          isRetrievable: true,
          labels: [
            'title'
          ]
        },
        {
          name: 'url',
          type: 'string',
          isRetrievable: true,
          labels: [
            'url'
          ]
        },
        {
          name: 'description',
          type: 'string',
          isRetrievable: true
        },
        {
          name: 'iconUrl',
          type: 'string',
          isRetrievable: true,
          labels: [
            'iconUrl'
          ]
        }
      ]
    }
  } as ExternalConnectors.ExternalConnection
};