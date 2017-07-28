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

let rgTemplates = new Array<Template>(10);
rgTemplates[0] = new Template("Sup Homies", "Foobar");
rgTemplates[1] = new Template("I'll Be Late", "I am running late and can't make it");
rgTemplates[2] = new Template("Sup Homies", "Foobar");
rgTemplates[3] = new Template("I'll Be Late", "I am running late and can't make it");
rgTemplates[4] = new Template("Sup Homies", "Foobar");
rgTemplates[5] = new Template("I'll Be Late", "I am running late and can't make it");
rgTemplates[6] = new Template("Sup Homies", "Foobar");
rgTemplates[7] = new Template("I'll Be Late", "I am running late and can't make it");
rgTemplates[8] = new Template("Sup Homies", "Foobar");
rgTemplates[9] = new Template("I'll Be Late", "I am running late and can't make it");

let myInfo : Info = new Info();

Office.initialize = () => {
    myInfo.updateName((Office.context.mailbox.item as Office.MessageRead).subject);
}

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
class ButtonBoard extends React.Component<ButtonBoardProps, undefined> {
    render() {
        return (
            <div className="buttonBoard">{this.props.buttons.map(button  => <SquareButton value={button.Title} />)}</div>
        )
    }
}

ReactDOM.render(
    (<ButtonBoard buttons={rgTemplates} />),
        document.getElementById("app")
);


/*
ReactDOM.render(
    (<SquareButton value="Robot Chicken Hello World"/>),
        document.getElementById("app")
        );
*/
