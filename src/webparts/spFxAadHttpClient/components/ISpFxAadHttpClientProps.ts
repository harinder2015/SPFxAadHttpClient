import {IUserItem} from '../Models/IUserItem'
import{MSGraphClient} from '@microsoft/sp-http'
export interface ISpFxAadHttpClientProps {
    userItems: IUserItem[];
    graphClient: MSGraphClient;
}
