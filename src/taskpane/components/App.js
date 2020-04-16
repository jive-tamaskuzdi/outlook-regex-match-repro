import * as React from "react";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
  }


  render() {
    const matches = Office.context.mailbox.item.getSelectedRegExMatches().mymatch;

    return <p>you have clicked: {JSON.stringify(matches)}</p>
  }
}
