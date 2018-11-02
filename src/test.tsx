import * as React from "react";
import * as ReactDOM from "react-dom";
import {action, observable, computed} from "mobx";
import {observer} from "mobx-react";
import { allowStateChangesStart } from "mobx/lib/core/action";
import {onClickInstantSend} from "../src/ews"

class ewsUpdate 
{
    // this is a hack to allow the callback from ewsRequest to set variables back into my ButtonBoard2 class
    @observable private _fSending:boolean = false;
    @computed public get FSending() : boolean {return this._fSending}
    @action setSending(inValue:boolean) : void {this._fSending = inValue;}
}

declare global {
    interface Window {
        uiLessHandler: any;
        timeOut: any; 
        ewsSend: ewsUpdate;
    }
}
window.ewsSend = new ewsUpdate;
window.ewsSend.setSending(false);

class Template {
    @observable private _title:string;
    @observable private _body:string;
    @observable private _id:number;
    @observable private _fDefault:boolean;

    constructor(inTitle:string, inBody:string, inId:number) {
        this.updateTitle(inTitle);
        this.updateBody(inBody);
        this.updateId(inId);

        this.updateFDefault(false);
    }

    @computed public get Title():string {return this._title;}
    @computed public get Body():string {return this._body;}
    @computed public get Id():number {return this._id;}
    @computed public get FDefault():boolean {return this._fDefault;}

    @action updateFDefault(inFDefault:boolean) : void {
        this._fDefault = inFDefault;
    }
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
    @observable private _fLoadedFromDisk:boolean = false;

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
    public addTemplate(inTitle:string, inBody:string) : number {
        this._rgTemplates.push(new Template(inTitle, inBody, this._currentIndex));
        if (this._currentIndex == 0)
            this._rgTemplates[0].updateFDefault(true); // Set the first template to be default
        this._currentIndex++;
        return (this._currentIndex - 1);
    }

    @action
    public deleteTemplate(inId:number)
    {
        for (var i = 0; i < this._rgTemplates.length; i++) {
            if (this._rgTemplates[i].Id == inId)
            {
                this._rgTemplates.splice(i, 1);
                break;
            }
        }
    }

    @action public changeDefault(inButtonId:number)
    {
        for (let i = 0; i < this._rgTemplates.length; i++ ) {
            this._rgTemplates[i].updateFDefault(this._rgTemplates[i].Id == inButtonId);
        }
  
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
            Default:boolean;
        };

        let myStructure : Array<tempTemplates> = new Array<tempTemplates>(0);

        for (let entry of this._rgTemplates )
        {
            let myEntry : tempTemplates = new tempTemplates;
            myEntry.Title = entry.Title;
            myEntry.Body = entry.Body;
            if (entry.FDefault)
                myEntry.Default = true;
            
            myStructure.push(myEntry);
        }

        return JSON.stringify(myStructure);
    }
}


class GlobalSettings
{
    @observable private _fReplyAll:boolean = false;
    @computed public get FReplyAll() : boolean {return this._fReplyAll}
    @action setReplyAll(inValue:boolean) : void {this._fReplyAll = inValue; this.saveToApplicationSettings();}

    @observable private _fEditResponse:boolean = true;
    @computed public get FEditResponse() : boolean {return this._fEditResponse; }
    @action setEditResponse(inValue:boolean) : void {this._fEditResponse = inValue;this.saveToApplicationSettings();}

    @action loadFromSettings()
    {
        var savedEditResponse = Office.context.roamingSettings.get("er2");
        var savedReplyAll = Office.context.roamingSettings.get("ra");

        if (savedEditResponse !== undefined)
            this._fEditResponse = savedEditResponse;
        
        if (savedReplyAll !== undefined)
            this._fReplyAll = savedReplyAll;
    }
    @action saveToApplicationSettings() 
    {
        Office.context.roamingSettings.set("er2", this._fEditResponse);
        Office.context.roamingSettings.set("ra", this._fReplyAll);
        Office.context.roamingSettings.saveAsync();
    }
}

let myTemplates : Templates = new Templates;
let myGlobalSettings : GlobalSettings = new GlobalSettings;

function saveToApplicationSettings(templatesToSave:Templates)
{
    let jsonString:string = templatesToSave.dumpJson();
 //   console.log(templatesToSave.dumpJson());
    Office.context.roamingSettings.set("templates", jsonString);
    Office.context.roamingSettings.saveAsync();
}

function UpdateTemplates()
{
    var savedSettings = Office.context.roamingSettings.get("templates"); 

   if ( savedSettings == undefined)
        {
            let tempTemplates = new Templates;
            tempTemplates.addTemplate("Default 2", "This is the <b>second</b> button.");
            tempTemplates.addTemplate("Default 1", "Body Text 1");
            tempTemplates.addTemplate("Default 3", "Body Text 1");
            tempTemplates.addTemplate("Default 4", "Body Text 1");
            tempTemplates.addTemplate("Default 5", "Body Text 1");
            tempTemplates.addTemplate("Default 6", "Body Text 1");
            myTemplates.updateTemplates(tempTemplates);
        }
    else
        {
            LoadTemplatesFromString(savedSettings);
        }
}

function LoadTemplatesFromString(stringIn:string)
{
    let jsonTemplates = JSON.parse(stringIn);
    let tempTemplates : Templates = new Templates;
    
    for (let i : number = 0; i < jsonTemplates.length; i++)
        {
            var id = tempTemplates.addTemplate(jsonTemplates[i].Title, jsonTemplates[i].Body);
            if (jsonTemplates[i].Default != undefined && jsonTemplates[i].Default == true)
            {
                tempTemplates.changeDefault(id);
            }
        }
    myTemplates.updateTemplates(tempTemplates);
}

LoadTemplatesFromString("[{\"Title\":\"Loading...\", \"Body\":\"Body\"}, {\"Title\":\"Loading...\", \"Body\":\"Body\"}]");

function ItemChanged(eventArgs:any)
{
    // Do nothing here.
}

Office.initialize = () => {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, ItemChanged);
    UpdateTemplates();
    myGlobalSettings.loadFromSettings();
}

export interface SquareButtonProps { value: string; onClick: any; onClickEdit: any;}
class SquareButton extends React.Component<SquareButtonProps, undefined > {
    render() {
        return(
            <span className="templateButtonHolder">
                <button className="templateButton" onClick={this.props.onClick}>{this.props.value}</button>
                <button className="editButton" onClick={this.props.onClickEdit}><img src = "icons/edit.png"></img></button>
            </span>
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

function saveReplyAllSetting(newSetting:boolean)
{
    Office.context.roamingSettings.set("replyall", newSetting);
    Office.context.roamingSettings.saveAsync();
}

export interface ButtonBoard2Props {inPageManager:PageManager;}
@observer
class ButtonBoard2 extends React.Component<ButtonBoard2Props, undefined> {
    
    constructor()
    {
        super();
    }

    quickReply(button:Template) {

    }

    // Called back when the ewsCall has finished for Instant Send
    ewsFinish()
    {
        window.ewsSend.setSending(false);
    }

    handleClick(button:Template) {
        if (myGlobalSettings.FEditResponse == false)
        {
            window.ewsSend.setSending(true);
            onClickInstantSend((Office.context.mailbox.item as Office.MessageRead).itemId, button.Body, myGlobalSettings.FReplyAll, this.ewsFinish);
        }
        else
        {
            if (myGlobalSettings.FReplyAll)
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
        myGlobalSettings.setReplyAll(!myGlobalSettings.FReplyAll);
    }

    handleEditResponseClick()
    {
        myGlobalSettings.setEditResponse(!myGlobalSettings.FEditResponse);
    }

    renderCheckbox(buttonText:string, checked:boolean, clickHandler:any){
        return <Checkbox2 onClick={clickHandler}checked={checked} text={buttonText}></Checkbox2>;
    }

    handleNewTemplate()
    {
        this.props.inPageManager.handleNewTemplate();
    }

    render() {
        if (window.ewsSend.FSending)  {
            return (<div className="loaderHolder"><div className="loader"></div>Sending...</div>)
        }
        else {
            return (
            <div className="buttonBoard"><div>{myTemplates.Data.map(button  => {
                var myString = button.Title;
                if (button.FDefault) 
                    myString += "*";
                return <SquareButton onClick={() => this.handleClick(button)} value={myString} onClickEdit={() => this.handleEditTemplateClick(button)} />
            })}</div>
            <div>{this.renderCheckbox("Reply All", myGlobalSettings.FReplyAll, () => this.handleReplyAllClick())}</div>
            <div>{this.renderCheckbox("Edit Response", myGlobalSettings.FEditResponse, () => this.handleEditResponseClick()) }</div> 
            <button onClick={() => this.handleNewTemplate()}className="newTemplateButton">Add New Template</button>
            </div>
            )
        }
    }
}

export interface EditTemplateState {body:string, title:string};
export interface EditTemplateFormProps {templateToEdit: Template; parentPageManager:PageManager;}
@observer
class EditTemplateForm extends React.Component<EditTemplateFormProps, EditTemplateState>
{
    constructor(props:EditTemplateFormProps) {
        super(props);
        if (this.props.templateToEdit == null)
        {
            this.state = {
                body: "Type Body Here",
                title: "New Template Name"
                };
        }
        else
        {
            this.state = {
            body: this.props.templateToEdit.Body,
            title: this.props.templateToEdit.Title
            };
        }        
        this.handleChange = this.handleChange.bind(this);
        this.handleSubmit = this.handleSubmit.bind(this);
        this.handleDiscard = this.handleDiscard.bind(this);
        this.handleDelete = this.handleDelete.bind(this);
        this.handleDefault = this.handleDefault.bind(this);
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

        if (this.props.templateToEdit == null)
        {
            myTemplates.addTemplate(newTitle, this.state.body);
        }
        else
        {
            myTemplates.changeTemplate(this.props.templateToEdit.Id, newTitle, this.state.body)
        }
        event.preventDefault();
        saveToApplicationSettings(myTemplates);
        this.props.parentPageManager.backToMain();
      }

      handleDiscard() {
        if (this.props.templateToEdit != null) {
            this.setState({
                body: this.props.templateToEdit.Body,
                title: this.props.templateToEdit.Title
            });
        }
          this.props.parentPageManager.backToMain();
      }

      handleDelete() {
        myTemplates.deleteTemplate(this.props.templateToEdit.Id);
        saveToApplicationSettings(myTemplates);
        this.props.parentPageManager.backToMain();
      }

      handleDefault() {
        myTemplates.changeDefault(this.props.templateToEdit.Id);
        saveToApplicationSettings(myTemplates);
      }

    render()
    {
        return  <div><form onSubmit={this.handleSubmit}>
                    <div><input className="editTemplateTitle" maxLength={20} name="title" value={this.state.title} onChange={this.handleChange}></input></div>
                    <div><textarea className="editTemplateBody" name="body" value={this.state.body} onChange={this.handleChange} /></div>
                    <div>
                        <input className="editTemplateButton" type="submit" value="Save" />
                        
                    </div>
                </form>
                <button className="editTemplateButton" onClick={this.handleDiscard} name="discard">Discard</button>
                { this.props.templateToEdit != null ? <button className="editTemplateButton" onClick ={this.handleDelete} name="delete">Delete</button> : null }
                { (this.props.templateToEdit == null) ? null : (this.props.templateToEdit.FDefault) ? 
                <button className="editTemplateButtonDefault" name = "default">Currently Default</button> : 
                <button className="editTemplateButton" onClick={this.handleDefault} name = "default">Make Default</button>}
                </div>
    }
}

@observer
class PageManager extends React.Component<{}, {}>
{
    @observable _fDisplayEdit : boolean = false;
    @observable _templateToEdit : Template = null;

    handleNewTemplate()
    {
        this._fDisplayEdit = true;
        this._templateToEdit = null;
    }

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
                <EditTemplateForm parentPageManager={this} templateToEdit={this._templateToEdit} />
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





window.uiLessHandler = function uiLessHandler(eventArgs:any)
{
    UpdateTemplates();
    myGlobalSettings.loadFromSettings();
    var defaultBody = "";

    for (var i = 0; i < myTemplates.Data.length; i++)
    {
        if (myTemplates.Data[i].FDefault){
            defaultBody = myTemplates.Data[i].Body;
            break;
        }
    }
    if (myGlobalSettings.FReplyAll == true) {
        (Office.context.mailbox.item as Office.MessageRead).displayReplyAllForm(defaultBody);
    }
    else {
        (Office.context.mailbox.item as Office.MessageRead).displayReplyForm(defaultBody);
    }

    setTimeout(function() {window.timeOut(eventArgs)}, 500);
}

window.timeOut = function timeOut(eventArgs:any)
{
    eventArgs.completed(true); 
}