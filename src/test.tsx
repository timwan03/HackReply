import * as React from "react";
import * as ReactDOM from "react-dom";
import {action, observable, computed} from "mobx";
import {observer} from "mobx-react";


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

    constructor(inTitle:string, inBody:string) {
        this.updateTitle(inTitle);
        this.updateBody(inBody);
    }

    @computed public get Title():string {return this._title;}

    @action updateTitle(inTitle:string) : void {
        this._title = inTitle;
    }

    @action updateBody(inBody:string) : void {
        this._body = inBody;
    }
}

class Templates {
    @observable private _rgTemplates:Array<Template>;

    @computed public get Data(): Array<Template>{return this._rgTemplates}

    @action
    public updateTemplates(inTemplates:Array<Template>) : void {
        this._rgTemplates = inTemplates;
    }
}
let rgTemplates = new Array<Template>(0);
rgTemplates.push(new Template("DebugTim", "Debug"));
let myTemplates : Templates = new Templates;

myTemplates.updateTemplates(rgTemplates);

let myInfo : Info = new Info();

function UpdateTemplates()
{
    var savedSettings = Office.context.roamingSettings.get("temp"); 

   if ( savedSettings == undefined)
        {
            let tempTemplates = new Array<Template>(0);
            tempTemplates.push(new Template("Default 2", "Edit Me"));
            tempTemplates.push(new Template("Default 1", "Edit Me"));
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
    let tempTemplates : Array<Template> = new Array<Template>(0);
    
    for (let i : number = 0; i < jsonTemplates.length; i++)
        {
            tempTemplates.push(new Template(jsonTemplates[i]._title, jsonTemplates[i]._body));

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

export interface SquareButtonProps { value: string;}
class SquareButton extends React.Component<SquareButtonProps, undefined > {
    
    render() {
        return(
            <button className="squareButton">{this.props.value}</button>
        )
    }
}

export interface ButtonBoardProps {buttons: Array<Template>}
@observer
class ButtonBoard extends React.Component<ButtonBoardProps, undefined> {
    render() {
        return (
            <div className="buttonBoard">{this.props.buttons.map(button  => <SquareButton value={button.Title} />)}</div>
        )
    }
}

@observer
class ButtonBoard2 extends React.Component<{}, {}> {
    render() {
        return (
            <div className="buttonBoard">{myTemplates.Data.map(button  => <SquareButton value={button.Title} />)}</div>
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
