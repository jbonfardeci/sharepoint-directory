import { Component, OnInit } from '@angular/core';
import { Http, Response, RequestOptions, Headers } from '@angular/http';
import { Observable } from 'rxjs/Rx';

declare var $: any; // JQuery

// SharePoint interfaces
interface IWrapper {
  d: ICollection;
}

interface ICollection {
  results: IPerson[];
}

interface IPerson {
  __metadata: IMetaData;
  Id: number;
  LastName: string;
  FirstName: string;
  EMail: string;
  Picture: string;
  Department: string;
  JobTitle: string;
  WorkPhone: string;
  Office: string;
}

interface IMetaData{
  id: string;
  uri: string;
  etag: string;
  type: string;
}

@Component({
  selector: 'app-directory',
  templateUrl: './directory.component.html',
  styleUrls: ['./directory.component.css']
})
export class DirectoryComponent implements OnInit {

  error: any;
  alpha: string[]; 
  people: IPerson[]; 
  person: IPerson; // selected person profile
  initial: string; // the current last name initial selected 

  // SharePoint API querystring parameters
  // select fields
  select = 'Id,LastName,FirstName,EMail,Picture,Department,JobTitle,WorkPhone,Office';
  // filter by people with last name initial eq initial
  filter = "startswith(LastName, '{init}') and ContentType eq 'Person' and LastName ne null'";
  orderby = 'LastName';
  url = "/_api/web/SiteUserInfoList/items";

  constructor(private http: Http){}

  ngOnInit(): void {
    this.alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  }

  setInitial(event, initial): void {
    var btn = event.target;
    this.initial = initial;
    this.person = undefined;
    this.getPeopleAlpha();
  }

  showDetail(person: IPerson): void {
    this.person = person;
  }

  showList(): void {
    this.person = undefined;
  }

  search(event): void {
    var self = this;
    this.error = undefined; 
    try{

      var term: string = event.target.value;
      var n: number = 3; // minimum number characters required to search

      // return results only when user has typed n characters in serach field
      if(term.replace(/\s+/, '').length < n) {
        return;
      }

      // search filter
      var filter = 
`(substringof('${term}', Title)
  or substringof('${term}', JobTitle)
  or substringof('${term}', WorkPhone)
  or substringof('${term}', EMail)
  or substringof('${term}', Office)
  or substringof('${term}', Department))
  and (ContentType eq 'Person' and LastName ne null)`;

      var url = `${this.url}?$select=${this.select}&$filter=${filter}&$orderby=${this.orderby}`;

      // local development mode
      if(window.location.hostname == 'localhost'){
        url = './assets/mockdata/users.json';
        this.http.get(url).subscribe((r: Response) => {
          var data = r.json();
          var people = data.d.results;
          var rx = new RegExp(term, 'gi');
          this.people = people.filter((p: IPerson): boolean => {
            return rx.test(p.FirstName + ' ' + p.LastName) 
              || rx.test(p.Department) 
              || rx.test(p.EMail) 
              || rx.test(p.JobTitle) 
              || rx.test(p.Office) 
              || rx.test(p.WorkPhone);
          });
        });
      }
      else {
        // production server mode from SharePoint 2013 API
        this.executeQuery(url);
      }
    }
    catch(e){
      this.error = e.toString();
    }
  }

  getPeopleAlpha(): void {
    var self = this; 
    this.error = undefined;  
    try {
      var email = '';
      var initial = this.initial || 'A';
      var filter = this.filter.replace('{init}', initial);

      var url = `${this.url}?$select=${this.select}&$filter=${filter}&$orderby=${this.orderby}`; 
      
      // local development mode
      if(window.location.hostname == 'localhost'){
        url = './assets/mockdata/users.json';
        this.http.get(url).subscribe((r: Response) => {
          var data = r.json();
          var people = data.d.results;
          var rx = new RegExp('^' + initial);
          self.people = people.filter((p: IPerson): boolean => {
            return rx.test(p.LastName);
          });
        });
      }
      else {
        // production server mode from SharePoint 2013 API
        this.executeQuery(url);
      }
    }
    catch(e){
      this.error = e.toString();
    }

  }

  executeQuery(url): void {
    var self = this;
    this.error = undefined; 
    var options: RequestOptions = new RequestOptions();
    options.headers = new Headers();
    options.headers.append('Accept', 'application/json;odata=verbose'); // header telling SharePoint API to return JSON, not XML

    this.http.get(encodeURI(url), options).subscribe((r: Response) => {
      if(r.ok){
        var data: IWrapper = r.json();
        self.people = data.d.results;
      }
      else {
        self.error = `An error occurred. Status: ${r.status}, ${r.statusText}`;
      }
    });
  }
}
