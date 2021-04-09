import * as React from 'react';
import styles from './SpFxAadHttpClient.module.scss';
import { ISpFxAadHttpClientProps } from './ISpFxAadHttpClientProps';
import { ISpFxAadHttpClientState } from './ISpFxAadHttpClientState';

import { escape } from '@microsoft/sp-lodash-subset';
import {Persona,personaSize,PersonaSize} from 'office-ui-fabric-react/lib/components/Persona';
import {Link} from 'office-ui-fabric-react/lib/components/Link';
import {MSGraphClient} from '@microsoft/sp-http'
export default class SpFxAadHttpClient extends React.Component<ISpFxAadHttpClientProps,ISpFxAadHttpClientState,{}> {

  constructor(props:ISpFxAadHttpClientProps)
  {
  super(props);
  this.state = {
 
   name:'',
   email:'',
   phone:'',
   image:null

  };

  }
  
  private _renderEmail =()=>
  {
if(this.state.email)
{
  return <Link href ={`mailto:${this.state.email}`}>{this.state.email}</Link>
}
else
return <div />;
  }
private _renderPhone =()=>
  {
    if(this.state.phone)
{
  return <Link href ={`tel:${this.state.phone}`}>{this.state.phone}</Link>
}
else
return <div />;
  }

  public componentDidMount():void{
    this.props.graphClient.api(`/me`).get((err:any,user: any)=>{
      //console.log("User :" + user.json().value)
     this.setState({
       name : user.displayName,
       email: user.mail,
       phone: user.businessPhones[0]
     })

    })
    this.props.graphClient.api(`/me/photo/$value`).responseType('blob').get((err:any,results:any)=>
    {
      const blobURL = window.URL.createObjectURL(results);
      this.setState({image:blobURL});
    });
    
  }
  public render(): React.ReactElement<ISpFxAadHttpClientProps> {
    return (
      <div className={ styles.spFxAadHttpClient }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>AadHttpClient AzureDemo</span>
              
            </div>
          </div>
          <div className={styles.row}>
         <div><strong> Mail: </strong></div>
         <ul className={styles.list}>
       {this.props.userItems.map((user)=>
     
     <li key={user.id} className={styles.item}>
       <strong>ID: </strong> {user.id}<br />
       <strong>Email:</strong> { user.mail }<br />
              <strong>DisplayName:</strong> { user.displayName }
        </li>
     
     )}
        
         </ul>

          </div>
<Persona 
text={this.state.name}
secondaryText={this.state.email}
onRenderSecondaryText={this._renderEmail}
tertiaryText ={this.state.phone}
onRenderTertiaryText={this._renderPhone}
imageUrl ={this.state.image}
size ={PersonaSize.size100} />


       </div>
      </div>
    );
  }
}
