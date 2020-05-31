import { Session,ITermSet,ITermData, ITerm } from "@pnp/sp-taxonomy";

export default class SPTaxonomyService {

    constructor() {}

    public async getTermSetTags(_siteCollectionUrl,_termStoreGuid,_termSetId){
        const _taxonomy = new Session(_siteCollectionUrl);
        const store = await _taxonomy.termStores.getById(_termStoreGuid).get();
        const termSet: ITermSet = store.getTermSetById(_termSetId);
        const termsWithData: (ITermData & ITerm)[] = await termSet.terms.select('Id', 'Name', 'Parent').get();
        return termsWithData;
    }

    public cleanGuid(guid: string): string {
        if (guid !== undefined) {
            return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
        } else {
            return '';
        }
    }

    public getTermSetTree = (
        data = [], 
        {idKey='id',parentKey='parentId',childrenKey='children'} = {}
    ) => {
        const tree = [];
        const childrenOf = {};
        data.forEach((item) => {
            const { [idKey]: id, [parentKey]: parentId = 0 } = item;
            childrenOf[id] = childrenOf[id] || [];
            item[childrenKey] = childrenOf[id];
            parentId 
              ? (
                  childrenOf[parentId] = childrenOf[parentId] || []
                ).push(item) 
              : tree.push(item);
        });
        return tree;
    }
}