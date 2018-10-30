import * as React from "react";
import * as ReactDOM from "react-dom";
import {action, observable, computed} from "mobx";
import {observer} from "mobx-react";
import { allowStateChangesStart } from "mobx/lib/core/action";


class Info {
    @observable private _subject:string;

    @computed public get Subject():string {return this._subject;}

    @action
    public updateName(inSubject:string) : void {
        this._subject = inSubject;
    }
}

class Template {
    @observable private _title:string;
    @observable private _body:string;
    @observable private _id:number;

    constructor(inTitle:string, inBody:string, inId:number) {
        this.updateTitle(inTitle);
        this.updateBody(inBody);
        this.updateId(inId);
    }

    @computed public get Title():string {return this._title;}
    @computed public get Id():number {return this._id;}
    @action updateTitle(inTitle:string) : void {
        this._title = inTitle;
    }

    @action updateBody(inBody:string) : void {
        this._body = inBody;
    }

    @action updateId(inId:number) : void {
        this._id = inId;
    }
}

class Templates {
    @observable private _rgTemplates:Array<Template>;
    @observable private _currentIndex:number;

    constructor() {
        this._currentIndex = 0;
        this._rgTemplates = new Array<Template>(0);
    }
    @computed public get Data(): Array<Template>{return this._rgTemplates}
    @computed public get CurrentIndex() : number {return this._currentIndex}

    @action
    public updateTemplates(inTemplates:Templates) : void {
        this._rgTemplates = inTemplates.Data;
        this._currentIndex = inTemplates.CurrentIndex;
    }

    @action
    public addTemplate(inTitle:string, inBody:string) {
        this._rgTemplates.push(new Template(inTitle, inBody, this._currentIndex));
        this._currentIndex++;
    }

    @action
    public changeTemplate(inButtonName:string) : void {
        for (let i = 0; i < this._rgTemplates.length; i++ ) {
            if (this._rgTemplates[i].Title == inButtonName)
            this._rgTemplates[i].updateTitle("hi mom");
        }
    }
}

//let rgTemplates = new Array<Template>(0);
//rgTemplates.push(new Template("DebugTim", "Debug"));
let myTemplates : Templates = new Templates;
myTemplates.addTemplate("DebugTim", "Debug");

let myInfo : Info = new Info();

function UpdateTemplates()
{
    var savedSettings = Office.context.roamingSettings.get("temp"); 

   if ( savedSettings == undefined)
        {
            let tempTemplates = new Templates;
            tempTemplates.addTemplate("Default 2", "Edit Me");
            tempTemplates.addTemplate("Default 1", "Edit Me");
            myTemplates.updateTemplates(tempTemplates);
        }
    else
        {


        }
}

function LoadTemplatesFromString()
{
    var stringIn = "[{\"_title\":\"LoadedFromDisk22\", \"_body\":\"Body\"}, {\"_title\":\"LoadedFromDisk\", \"_body\":\"Body\"}]";
    let jsonTemplates = JSON.parse(stringIn);
    let tempTemplates : Templates = new Templates;
    
    for (let i : number = 0; i < jsonTemplates.length; i++)
        {
            tempTemplates.addTemplate(jsonTemplates[i]._title, jsonTemplates[i]._body);
        }
    myTemplates.updateTemplates(tempTemplates);
}

LoadTemplatesFromString();

Office.initialize = () => {
    //myInfo.updateName((Office.context.mailbox.item as Office.MessageRead).subject);
    UpdateTemplates();
}

//setTimeout(UpdateTemplates, 1000);

@observer
class HelloWorld extends React.Component<{}, {}> {
    public render(): JSX.Element {
        return <div> This Message's Subject is: {myInfo.Subject}</div>;
    }
}

export interface SquareButtonProps { value: string; onClick: any;}
class SquareButton extends React.Component<SquareButtonProps, undefined > {
    render() {
        return(
            <button className="squareButton" onClick={this.props.onClick}>{this.props.value}</button>
        )
    }
}
/*
export interface ButtonBoardProps {buttons: Array<Template>}
@observer
class ButtonBoard extends React.Component<ButtonBoardProps, undefined> {
    render() {
        return (
            <div className="buttonBoard">{this.props.buttons.map(button  => <SquareButton value={button.Title} />)}</div>
        )
    }
}
*/

@observer
class ButtonBoard2 extends React.Component<{}, {}> {

    handleClick(button:Template) {
        // window.open("https://www.yahoo.com/"); // TODO: Remove        
        // myTemplates.changeTemplate("Default 2");
    }

    render() {
        return (
            <div className="buttonBoard">{myTemplates.Data.map(button  => <SquareButton onClick={() => this.handleClick(button)} value={button.Title} />)}</div>
        )
    }
}


ReactDOM.render(
    (<ButtonBoard2  />),
        document.getElementById("app")
);

/*
ReactDOM.render(
    (<ButtonBoard buttons={myTemplates.Data} />),
        document.getElementById("app")
);


/*
ReactDOM.render(
    (<SquareButton value="Robot Chicken Hello World"/>),
        document.getElementById("app")
        );
*/
