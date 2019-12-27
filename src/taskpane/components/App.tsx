import * as React from "react";
// import { Button, ButtonType } from "office-ui-fabric-react";
import Progress from "./Progress";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */
import SheetList from "./SheetList/SheetList";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export default class App extends React.Component<AppProps> {
  constructor(props, context) {
    super(props, context);
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <SheetList />
    );
  }
}
