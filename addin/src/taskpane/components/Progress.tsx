import React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";

export type ProgressProps = {
  logo: string;
  message: string;
  title: string;
}

const Progress = (props: ProgressProps) => {
  return (
    <section className="ms-welcome__progress ms-u-fadeIn500">
      <img width="90" height="90" src={props.logo} alt={props.title} title={props.title} />
      <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{props.title}</h1>
      <Spinner size={SpinnerSize.large} label={props.message} />
    </section>
  );
}

export default Progress;