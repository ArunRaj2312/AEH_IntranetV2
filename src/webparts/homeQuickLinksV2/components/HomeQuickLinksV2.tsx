import * as React from "react";
import styles from "./HomeQuickLinksV2.module.scss";
import type { IHomeQuickLinksV2Props } from "./IHomeQuickLinksV2Props";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import MainComponent from "./MainComponent";

export default class HomeQuickLinksV2 extends React.Component<IHomeQuickLinksV2Props> {
  public render(): React.ReactElement<IHomeQuickLinksV2Props> {
    return (
      <MainComponent
        sp={spfi().using(SPFx(this.props.context))}
        context={this.props.context}
        listTitle="IntranetComponents"
      />
    );
  }
}
