import * as React from "react";
import * as ReactDOM from "react-dom";
import {action, observable, computed} from "mobx";
import {observer} from "mobx-react";
import { allowStateChangesStart } from "mobx/lib/core/action";
import {onClickInstantSend} from "../src/ews"

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
    @computed public get Body():string {return this._body;}
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
    public changeTemplate(inButtonId:number, inTitle:string, inBody:string) : void {
        for (let i = 0; i < this._rgTemplates.length; i++ ) {
            if (this._rgTemplates[i].Id == inButtonId)
            {
                if (inTitle)
                    this._rgTemplates[i].updateTitle(inTitle);
                if (inBody)
                    this._rgTemplates[i].updateBody(inBody);
            }
        }
    }

    @action
    public dumpJson() : string {
        class tempTemplates{
            Title:string;
            Body:string;
        };

        let myStructure : Array<tempTemplates> = new Array<tempTemplates>(0);

        for (let entry of this._rgTemplates )
        {
            let myEntry : tempTemplates = new tempTemplates;
            myEntry.Title = entry.Title;
            myEntry.Body = entry.Body;
            
            myStructure.push(myEntry);
        }

        return JSON.stringify(myStructure);
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
            tempTemplates.addTemplate("Default 2 dasdfasf asdf asdf asdas fas fasd ", "This is the <b>second</b> button.");
            tempTemplates.addTemplate("Default 1", "Body Text 1");
            myTemplates.updateTemplates(tempTemplates);
        }
    else
        {


        }
}

function LoadTemplatesFromString(stringIn:string)
{
    let jsonTemplates = JSON.parse(stringIn);
    let tempTemplates : Templates = new Templates;
    
    for (let i : number = 0; i < jsonTemplates.length; i++)
        {
            tempTemplates.addTemplate(jsonTemplates[i].Title, jsonTemplates[i].Body);
        }
    myTemplates.updateTemplates(tempTemplates);
}

LoadTemplatesFromString("[{\"Title\":\"LoadedFromDisk22\", \"Body\":\"Body\"}, {\"Title\":\"LoadedFromDisk\", \"Body\":\"Body\"}]");

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

export interface SquareButtonProps { value: string; onClick: any; onClickEdit: any;}
class SquareButton extends React.Component<SquareButtonProps, undefined > {
    render() {
        return(
            <div>
                <button className="templateButton" onClick={this.props.onClick}>{this.props.value}</button>
                <button className="editButton" onClick={this.props.onClickEdit}><img src = "icons/edit.png"></img></button>
            </div>
        );
    }
}

export interface Checkbox2Props {checked:boolean, text:string, onClick: any;}
class Checkbox2 extends React.Component<Checkbox2Props, undefined>
{
    render() {
        let buttonString:string = " ";
        if (this.props.checked)
        {
            buttonString = "x";
        }
        return (
            <div className="checkBoxContainer">
             <button className="checkBox" onClick={this.props.onClick}>{buttonString}</button>
                {this.props.text}
            </div>
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

export interface ButtonBoard2Props {inPageManager:PageManager;}
@observer
class ButtonBoard2 extends React.Component<ButtonBoard2Props, undefined> {
    @observable _isReplyAll:boolean;
    @observable _editResponse:boolean
    constructor()
    {
        super();
        this._isReplyAll = false;
        this._editResponse = true;
    }

    quickReply(button:Template) {

    }

    handleClick(button:Template) {
        //myTemplates.changeTemplate(button.Id);
        console.log(myTemplates.dumpJson());

        if (this._editResponse == false)
        {
            onClickInstantSend((Office.context.mailbox.item as Office.MessageRead).itemId, button.Body, this._isReplyAll);
        }
        else
        {
        if (this._isReplyAll)
            (Office.context.mailbox.item as Office.MessageRead).displayReplyAllForm(button.Body)
        else
            (Office.context.mailbox.item as Office.MessageRead).displayReplyForm(button.Body)
        }

    }

    handleEditTemplateClick(button:Template) {
        this.props.inPageManager.handleEditClick(button);
    }

    handleReplyAllClick()
    {
        this._isReplyAll = !this._isReplyAll;
    }

    handleEditResponseClick()
    {
        this._editResponse = !this._editResponse;
    }

    renderCheckbox(buttonText:string, checked:boolean, clickHandler:any){
        return <Checkbox2 onClick={clickHandler}checked={checked} text={buttonText}></Checkbox2>;
    }

    render() {
        return (
        <div className="buttonBoard"><div>{myTemplates.Data.map(button  => {
            var myString = button.Title + ":" + button.Id;
            return <SquareButton onClick={() => this.handleClick(button)} value={myString} onClickEdit={() => this.handleEditTemplateClick(button)} />
        })}</div>
        <div>{this.renderCheckbox("Reply All", this._isReplyAll, () => this.handleReplyAllClick())}</div>
        <div>{this.renderCheckbox("Edit Response", this._editResponse, () => this.handleEditResponseClick()) }</div> 
        </div>
        )
    }
}

export interface EditTemplateState {body:string, title:string};
export interface EditTemplateFormProps {templateToEdit: Template; newTemplate:boolean; parentPageManager:PageManager;}
@observer
class EditTemplateForm extends React.Component<EditTemplateFormProps, EditTemplateState>
{
    constructor(props:EditTemplateFormProps) {
        super(props);
        this.state = {
          body: this.props.templateToEdit.Body,
          title: this.props.templateToEdit.Title
        };
        
        this.handleChange = this.handleChange.bind(this);
        this.handleSubmit = this.handleSubmit.bind(this);
      }

      handleChange(event:any) {
        const target = event.target;
        //var value = target === 
        if (target.name === "body") {
            this.setState({body: target.value as string});
        }
        else {


            this.setState({title: target.value as string});
        }
      }
    
      handleSubmit(event:any) {
        let newTitle:string = this.state.title.trim();
        if (newTitle.length == 0)
            newTitle = "<empty title>";
        //alert('A name was submitted: ' + this.state.value);
        myTemplates.changeTemplate(this.props.templateToEdit.Id, newTitle, this.state.body)
        event.preventDefault();
        this.props.parentPageManager.backToMain();
      }

    render()
    {
        return  <form onSubmit={this.handleSubmit}>
                    <input name="title" value={this.state.title} onChange={this.handleChange}></input>
                    <textarea name="body" value={this.state.body} onChange={this.handleChange} />
                    <input type="submit" value="Submit" />
                </form>
    }
}

@observer
class PageManager extends React.Component<{}, {}>
{
    @observable _fDisplayEdit : boolean = false;
    @observable _templateToEdit : Template;

    handleEditClick(button:Template)
    {
        this._fDisplayEdit = true;
        this._templateToEdit = button;
    }

    backToMain()
    {
        this._fDisplayEdit = false;
        this._templateToEdit = null;
    }

    render()
    {
        if (this._fDisplayEdit)
        {
            return (
                <EditTemplateForm parentPageManager={this} templateToEdit={this._templateToEdit} newTemplate={false} />
            )

        }

        return(
            <ButtonBoard2 inPageManager={this}/>
        )
    }
}
ReactDOM.render(
    (<PageManager  />),
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
