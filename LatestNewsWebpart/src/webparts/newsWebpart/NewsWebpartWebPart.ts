import { SPComponentLoader } from '@microsoft/sp-loader';
import * as pnp from 'sp-pnp-js';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewsWebpartWebPart.module.scss';
import * as strings from 'NewsWebpartWebPartStrings';

export interface INewsWebpartWebPartProps {
  description: string;
}
require('./app/style.css');
export default class NewsWebpartWebPart extends BaseClientSideWebPart<INewsWebpartWebPartProps> {
  private weburl:string

  public constructor() {
    super();
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js', { globalExportsName: 'jQuery' }).then((jQuery: any): void => {
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js',  { globalExportsName: 'jQuery' }).then((): void => {        
      });
    });
  }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      pnp.setup({
        spfxContext: this.context
      });
      
    });
  }

  public getDataFromList():void {
    var mythis =this;
    pnp.sp.web.lists.getByTitle('News').items.top(4).get().then(function(result){
      //console.log("Got NEws List Data:"+JSON.stringify(result));
      mythis.displayData(result);
    },function(er){
      alert("Oops, Something went wrong, Please try after sometime");
      console.log("Error:"+er);
    });


  }

  public displayData(data):void{
    var weburl = this.context.pageContext.web.absoluteUrl
    var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    data.forEach(function(val){
      var dt=new Date(val.Created);
      var title=val.Title;
      var subtitle=val.uvyk;
      var ID=val.ID;
      subtitle = subtitle.replace(/(<([^>]+)>)/g, "");
      if(title.length > 30){
        title = title.substring(0,30)+"...";
      }
      if(subtitle.length > 80){
        subtitle = subtitle.substring(0,80)+"...";
      }
      var myHtml = '<div class="col-sm-3">'+
        '    <div class="post-module">'+
        '       <a href='+weburl+'/Lists/News/DispForm.aspx?ID='+ID+'&Source='+weburl+'/SitePages/Intranet-Web-parts.aspx target="_blank">'+
        '        <div class="thumbnail">'+
        '            <div class="datee">'+
        '                <div class="day">'+dt.getDate()+'</div>'+
        '                <div class="month">'+monthNames[dt.getMonth()]+'</div>'+
        '            </div>'+
        '        </div>'+
        '        <div class="post-content">'+
        '            <div class="category1">'+val.tuog+'</div>'+
        '            <h1 class="title">'+title+'</h1>'+
        '            <h2 class="sub_title1">'+subtitle+'</h2>'+
        '            <p class="description" style="height:60px!important; overflow:hidden!important; ">"+Display_desc+"</p>'+
        '        </div>'+
        '       </a>'+
        '    </div>'+
        '</div>';

        var div = document.getElementById("newsWebpart");
        div.innerHTML+=myHtml;
    });
    
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="row">
	<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
		<div class="card card-stats events-news">
			<div class="card-header crdhead">News</div>
			<div class="card-content panel-body" id="newsWebpart">
				
			</div>
		</div>
		<div class="panel-footer" style="text-align:center">
			<a class="panelfooterA" href="/sites/Intranet/SPFX/Lists/News/AllItems.aspx" target="_blank">READ MORE</a>
		</div>
	</div>
</div>`;
this.weburl = this.context.pageContext.web.absoluteUrl;
console.log("------------"+this.context.pageContext.web.absoluteUrl);
this.getDataFromList();

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
