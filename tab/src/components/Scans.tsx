import {useState, useEffect, useContext } from "react";
import { useTeams } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "./Context";
import { makeStyles, Divider, MessageBar, MessageBarTitle, MessageBarBody, MessageBarActions, FluentProvider, teamsLightTheme, tokens, Table, TableHeader, TableRow, TableHeaderCell, TableBody, TableCell, Card, CardHeader, CardPreview, Text, Dropdown, Option, TableCellLayout, Button, Spinner, Dialog, DialogTrigger, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions, MenuItem, Field, RadioGroup, Radio, Body1, Caption1, CounterBadge, Badge, Subtitle2, Avatar } from "@fluentui/react-components";
import { MoreHorizontal20Regular } from "@fluentui/react-icons";
import * as microsoftTeams from "@microsoft/teams-js";
import { DismissRegular } from "@fluentui/react-icons";

import { Duplicate } from "./models/duplicate";
import { Scan } from "./models/scan";
import File from "./models/file";
import Config from "./Config";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    padding: "16px",
    gap: "16px",
  },
  header: {
    width: "100%",
    padding: "4px"
  },
});

export default function Scans() {
  const { theme } = useTeams({})[0];
  const [token, setToken] = useState<string>("");
  const [data, setData] = useState<Scan[]>();
  const [selectedScan, setSelectedScan] = useState<Scan>();
  const [waiting, setWaiting] = useState<boolean>(true);
  const [saving, setSaving] = useState<boolean>(false);
  const [scanning, setScanning] = useState<boolean>(false);
  const [refreshing, setRefreshing] = useState<boolean>(false);
  const [activeScan, setActiveScan] = useState<boolean>(false);
  const [errorText, setErrorText] = useState<string>("");
  const [fileToKeep, setFileToKeep] = useState<string>("");
  const [resolveIndex, setResolveIndex] = useState<number>(-1);
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  

  // useEffect to handle first load
  useEffect(() => {
    const firstLoad = async function () {
      const token = await teamsUserCredential?.getToken("");
      setToken(token?.token as string);
      await refreshData(token?.token);
    };
    firstLoad();
  }, []);

  // refreshes the scans data
  const refreshData = async (token?:string) => {
    setRefreshing(true);
    const scans = await fetch(`${Config.botEndpoint}/api/scans`, {
      method: "GET",
      headers: {
          "Authorization": `Bearer ${token}`,
          "Content-Type": "application/json",
      },
    });

    // parse the json and save state
    if (scans.ok) {
      var scansJSON = await scans.json();
      setData(scansJSON);
      setSelectedScan(scansJSON.find((i:any) => i.status === "complete"));
      setActiveScan(scansJSON.find((i:any) => i.status === "active"));
      setErrorText("");
    }
    else {
      setErrorText("Failed to query scans");
    }
    setRefreshing(false);
    setScanning(false);
    setWaiting(false);
  };

  // sets the resolve duplicate index
  const resolve = async (index:number) => {
    setResolveIndex(index);
  };

  // processes the resolution of a selected duplicate
  const processResolveSelection = async (saveChanges:boolean) => {
    if (saveChanges) {
      // set saving indicator
      setSaving(true);
      
      // prepare the delete
      const duplicate:Duplicate = selectedScan!.duplicates[resolveIndex];
      duplicate.fileToKeep = fileToKeep;
      await fetch(`${Config.botEndpoint}/api/files/${selectedScan?.id}`, {
        method: "DELETE",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json",
        },
        body: JSON.stringify(duplicate)
      });

      // remove the duplicate from the active scan
      setSelectedScan(prevScan => {
        if (!prevScan) return prevScan;
        return {
          ...prevScan, duplicates: prevScan.duplicates.filter((_, index) => index !== resolveIndex),
        } as Scan;
      });

      setSaving(false);
    }
    else {
      // increment index since this is just a skip
      setResolveIndex(resolveIndex+1);
    }
    
    // reset the file to keep state
    setFileToKeep("");
  };

  // starts a new scan
  const startScan = async () => {
    setScanning(true);
    const response = await fetch(`${Config.botEndpoint}/api/scans`, {
      method: "POST",
      headers: {
          "Authorization": `Bearer ${token}`,
          "Content-Type": "application/json",
      }
    });
    const scan:Scan = await response.json();
    setData((prevData) => [...(prevData|| []), scan]);
    setScanning(false);
    setActiveScan(true);
  };

  let cards:any = [];
  let options:any = [];
  const styles = useStyles();

  if (selectedScan) {
    selectedScan.duplicates.forEach((value:Duplicate, index:number) => {
      let locations = value.locations.map((v:File, i:number) => (
        <div key={v.id}>{v.path}</div>
      ));

      let radios = value.locations.map((v:File, i:number) => (
        <Radio value={v.path} label={v.path} />
      ));

      cards.push(
        <div style={{width: "100%"}}>
          <CardHeader 
            onClick={() => { resolve(index) }}
            className={styles.header}
            // TODO: this doesn't work for all mime types...need a better collection of doc icons
            image={{ as: "img", src: `https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/src/assets/${value.fileExt.substring(1)}.png`, alt: "Microsoft PowerPoint logo",}}
            header={<Body1><b>{value.fileName}</b></Body1>}
            description={value.size + " bytes"}
            action={<Badge appearance="filled">{value.locations.length + " copies"}</Badge>}
            style={{cursor: "pointer"}}
          />
          <Divider />
      </div>
      );
    });
    if (selectedScan.duplicates.length > 0)
      cards.push(<Divider />);
  };

  if (data && data.length > 0) {
    data.forEach((value:Scan, index: number) => {
      options.push(<Option>{value.scanDate}</Option>)
    })
  }

  // setup button icons
  const confirmIcon = (saving) ? (<Spinner size="tiny" />) : null;
  const scanIcon = (scanning) ? (<Spinner size="tiny" />) : null;
  const refreshIcon = (refreshing) ? (<Spinner size="tiny" />) : null;

  // setup banner to indicate if an active scan is in progress
  const banner = (activeScan) ? (
    <MessageBar style={{width: "100%", marginBottom: "10px"}}>
      <MessageBarBody>
        <MessageBarTitle>Scan in process</MessageBarTitle>
        Refresh to get updated results.
      </MessageBarBody>
      <MessageBarActions>
        <Button icon={refreshIcon} onClick={() => { refreshData(token)}}>Refresh</Button>
      </MessageBarActions>
    </MessageBar>
  ) : (<></>);

  // setup error bar
  const error = (errorText !== "") ? (
    <MessageBar intent="error" style={{width: "100%", marginBottom: "10px"}}>
      <MessageBarBody>
        <MessageBarTitle>Error occurred</MessageBarTitle>
        {"error"}
      </MessageBarBody>
      <MessageBarActions
      containerAction={
        <Button
          aria-label="dismiss"
          appearance="transparent"
          icon={<DismissRegular />}
          onClick={() => { setErrorText("") }}
        />}>
          <Button icon={refreshIcon} onClick={() => { refreshData(token)}}>Refresh</Button>
      </MessageBarActions>
    </MessageBar>
  ) : (<></>);

  const dialog = (resolveIndex !== -1) ? (
    <Dialog open={true}>
      <DialogSurface style={{maxWidth: "95vw"}}>
        <DialogBody>
          <DialogTitle style={{display: "flex", justifyContent: "space-between"}}>
            {
            // TODO: this doesn't work for all mime types...need a better collection of doc icons
            }
            <div><Avatar shape="square" style={{background: "transparent"}} image={{src: `https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/stories/src/assets/${selectedScan?.duplicates[resolveIndex].fileExt.substring(1)}.png`, className: "avatar-image"}} /> {selectedScan?.duplicates[resolveIndex].fileName}</div>
            <div>({resolveIndex + 1} of {selectedScan?.duplicates.length})</div>
          </DialogTitle>
          <DialogContent>
          <Field label="Select which file to keep:">
            <RadioGroup onChange={(e:any) => { setFileToKeep(e.target.value); }}>
              {
                selectedScan?.duplicates[resolveIndex].locations.map((v:File, i:number) => (
                  <Radio value={v.path} label={v.path} checked={fileToKeep === v.path} />
                ))
              }
            </RadioGroup>
          </Field>
          </DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary" onClick={() => { setResolveIndex(-1); }}>Cancel</Button>
            </DialogTrigger>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary" onClick={() => { processResolveSelection(false); }}>Skip</Button>
            </DialogTrigger>
            <DialogTrigger disableButtonEnhancement>
              <Button icon={confirmIcon} appearance="primary" disabled={(fileToKeep === "" || saving)} onClick={() => { processResolveSelection(true); }}>Confirm</Button>
            </DialogTrigger>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  ) : (<></>);

  return (
    <FluentProvider theme={ theme || { ...teamsLightTheme, colorNeutralBackground3: "#eeeeee" }}
      style={{ background: tokens.colorNeutralBackground3, width: "100%" }}>
        <div style={{display: "flex", flexDirection: "column", height: "100vh", width: "100%", justifyContent: "flex-start", alignItems: "flex-start", padding: "20px"}}>
          {(waiting) ? (<Spinner style={{margin: "auto"}} />) : (<></>)}
          {banner}
          {error}
          <div style={{display: (data) ? "flex" : "none", justifyContent: "space-between", paddingBottom: "10px", width: "100%"}}>
            <Subtitle2 style={{display: (selectedScan) ? "flex" : "none", marginTop: "4px"}}>Showing scan results as of {(new Date(selectedScan?.scanDate as string)).toLocaleString("en-us")}</Subtitle2>
            <Subtitle2 style={{display: (selectedScan) ? "none" : "flex",marginTop: "4px"}}>You have no scan history</Subtitle2>
            <Button icon={scanIcon} appearance="primary" disabled={activeScan || scanning} onClick={() => { startScan() }}>Start new scan</Button>
          </div>
          <div style={{width: "100%"}}>
            {cards}
            {dialog}
          </div>
          
        </div>
        <div style={{display: "flex", justifyContent: "flex-end", padding: "10px 20px", position: "sticky", bottom: "0"}}>
          <Button onClick={() => { 
            microsoftTeams.dialog.url.submit({ status: "completed" });
            return true;
            //microsoftTeams.tasks.submitTask({ status: "completed" });
          }}>Close</Button>
        </div>
    </FluentProvider>
  );
}