import * as React from 'react';
import styles from './TagCloud.module.scss';
import { ITagCloudProps } from './ITagCloudProps';
import { ITagCloudState } from './ITagCloudState';
import { escape } from '@microsoft/sp-lodash-subset';
import { TagCloud } from "react-tagcloud";
import { Web, CamlQuery } from '@pnp/sp';
import * as CamlBuilder from 'camljs';
import * as _ from 'underscore';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export default class SPTagCloud extends React.Component<ITagCloudProps, ITagCloudState> {
  private _web : Web;
  private _listName : string = 'Pages';
  private _data = [];
  
  private options = {
    hue: 'green',
    luminosity : 'dark'
  };

  private _customRenderer = (tag, size, color) => (
    <span key={tag.value}
          style={{
            fontSize: `${size}px`,
            border: `2px solid ${color}`,
            backgroundColor : `${color}`,
            margin: '3px',
            padding: '5px',
            display: 'inline-block',
            color: 'white',
            borderRadius: '10px',
            boxShadow : '2px 2px 1px rgba(128, 128, 128, 0.62)',
            cursor : 'pointer'
          }}><Icon iconName="Tag" /> {tag.value}</span>
  )
  constructor(props) {
    super(props);
    this.state = {
        data : []
    };
  }

  public render(): React.ReactElement<ITagCloudProps> {
    return (
      <div className={ styles.tagCloud }>
        <div className={ styles.container }>
          <div className={ styles.row }>
          <TagCloud minSize={8}
            maxSize={24}
            colorOptions={this.options}
            tags={this.state.data}
            renderer={this._customRenderer}
            onClick={tag => this._openSearch(tag)} />
          </div>
        </div>
      </div>
    );
  }
  private async _buildCAMLQuery () {
    let xml = new CamlBuilder().View(["W365_RelatedTopic"])
      .Query().Where().ModStatField("_ModerationStatus").IsApproved().And()
      .LookupMultiField("W365_RelatedTopic").IsNotNull().ToString();
      
      let query = CamlBuilder.FromXml(xml).ModifyWhere().AppendAnd()
        .TextField("ContentType").EqualTo("Wizdom Newspage").Or()
        .TextField("ContentType").EqualTo("Syngenta Positions Article");
      
    return query.ToString();
}

  public _openSearch = (tag)=>{
    let siteURL = this.props.context.pageContext.site.absoluteUrl;
    let searchURL = `${siteURL}/_layouts/15/osssearchresults.aspx?u=${siteURL}&k=owstaxIdW365RelatedTopic:${tag.value}`;
    window.open(encodeURI(searchURL), "_blank");
  }

private async _getTermsWithCAML(web: Web, listTitle: string) {
  return new Promise(async(resolve, reject) => {

      const items = web.lists.getByTitle(listTitle).items;

      let camlQuery = await this._buildCAMLQuery();
      const caml: CamlQuery = {
        ViewXml: camlQuery,
      };

      web.lists.getByTitle(listTitle).getItemsByCAMLQuery(caml,"W365_RelatedTopic").then(async i =>{
        resolve(i);
      }, fail =>{
         console.log(fail);
        reject(fail);
      });
  });
}
  private async _getTerms() {
    let webUrl : string = this.props.context.pageContext.web.absoluteUrl + '/articles';
    this._web = new Web(webUrl);
    
      try {
          let items = await this._getTermsWithCAML(this._web, this._listName);
          let string = JSON.stringify(items);
          let arr : any[] = JSON.parse(string);
          var list = [];
          _.each(arr, (item)=>{
              list.push(_.pick(item, 'W365_RelatedTopic'));
          });
          list = _(list).chain().zip(_(list).pluck('W365_RelatedTopic'))
            .flatten().compact().filter((item) =>{
              return !_.has(item, 'W365_RelatedTopic');
            }).value();
          let count=  _.countBy(list, "Label");
           _.mapObject(count, (val, key) =>{
              this._data.push({
                value: key,
                count : val
              });
           });
           this.setState({
             data : this._data
           });
      } catch (e) {
          console.error(e);
      }
  }

  public componentDidMount(){
    this._getTerms();
  }
}
