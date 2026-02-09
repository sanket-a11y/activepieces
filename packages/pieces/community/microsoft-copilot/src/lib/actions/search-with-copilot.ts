import { createAction, Property } from '@activepieces/pieces-framework';
import { microsoft365CopilotAuth } from '../common/auth';
import { httpClient, HttpMethod } from '@activepieces/pieces-common';

export const searchWithCopilot = createAction({
  auth: microsoft365CopilotAuth,
  name: 'searchWithCopilot',
  displayName: 'Search with Copilot',
  description:
    'Perform hybrid (semantic and lexical) search across OneDrive for work or school content using natural language queries',
  props: {
    query: Property.LongText({
      displayName: 'Search Query',
      description:
        'Natural language query to search for relevant files. Maximum 1,500 characters.',
      required: true,
    }),
    pageSize: Property.Number({
      displayName: 'Page Size',
      description: 'Number of results to return per page (1-100). Default: 25.',
      required: false,
      defaultValue: 25,
    }),
    filterExpression: Property.LongText({
      displayName: 'Filter Expression (Optional)',
      description:
        'KQL filter expression for OneDrive paths. Example: path:"https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Finance/"',
      required: false,
    }),
    resourceMetadata: Property.Array({
      displayName: 'Resource Metadata Names (Optional)',
      description: 'Array of metadata field names to include in results (e.g., title, author)',
      required: false,
    }),
  },
  async run(context) {
    const {
      query,
      pageSize,
      filterExpression,
      resourceMetadata,
    } = context.propsValue;

    const dataSources: any = {};

    if (filterExpression || (resourceMetadata && resourceMetadata.length > 0)) {
      dataSources.oneDrive = {};
      if (filterExpression) {
        dataSources.oneDrive.filterExpression = filterExpression;
      }
      if (resourceMetadata && resourceMetadata.length > 0) {
        dataSources.oneDrive.resourceMetadataNames = resourceMetadata;
      }
    }

    const body: {
      query: string;
      pageSize?: number;
      dataSources?: any;
    } = {
      query,
    };

    if (pageSize) {
      body.pageSize = pageSize;
    }

    if (Object.keys(dataSources).length > 0) {
      body.dataSources = dataSources;
    }

    const response = await httpClient.sendRequest({
      method: HttpMethod.POST,
      url: 'https://graph.microsoft.com/beta/copilot/search',
      headers: {
        'Authorization': `Bearer ${context.auth.access_token}`,
        'Content-Type': 'application/json',
      },
      body: body,
    });

    return response.body;
  },
});
