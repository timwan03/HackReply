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

ReactDOM.render(
    (<HelloWorld />),
        document.getElementById("app")
        );