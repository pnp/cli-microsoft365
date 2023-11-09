
interface GraphQueryParameters {
  /**
   * List of properties separated by a comma. Properties without a slash are used in $select query parameter. 
   * Propeties with a slash are used in $expand query parameter.
   */
  properties?: string;

  /** Filter expression used in $filter query parameter.*/
  filter?: string;

  /** If specified then $count=true is included.*/
  count?: boolean;
}

export const queryUtils = {
  /**
   * Create a query for a request to the Graph API
   * @param parameters Parameters to be applied to the query.
   * @returns Query with applied parameters.
  */
  createGraphQuery(parameters: GraphQueryParameters): string {
    const queryParameters: string[] = [];
    if (parameters.properties) {
      const allProperties = parameters.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));
      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }

      const expandProperties = allProperties.filter(prop => prop.includes('/'));

      let fieldExpand: string = '';
      expandProperties.forEach(p => {
        if (fieldExpand.length > 0) {
          fieldExpand += ',';
        }

        fieldExpand += `${p.split('/')[0]}($select=${p.split('/')[1]})`;
      });
      if (fieldExpand.length > 0) {
        queryParameters.push(`$expand=${fieldExpand}`);
      }
    }
    if (parameters.filter) {
      queryParameters.push(`$filter=${parameters.filter}`);
    }
    if (parameters.count) {
      queryParameters.push('$count=true');
    }

    let query = '';
    for (let i = 0; i < queryParameters.length; i++) {
      query += i === 0 ? '?' : '&';
      query += queryParameters[i];
    }
    return query;
  }
};