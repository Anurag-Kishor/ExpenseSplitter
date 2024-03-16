import { Injectable } from '@angular/core';
import { Group } from './group';

@Injectable({
  providedIn: 'root'
})
export class GroupService {

  constructor() { }
  groupsList : Group[] = []

  createGroup(group: Group) {
    console.log(`Group application received: Name: ${group.groupName}.`)
    this.groupsList.push(group)
    console.log(this.groupsList)
  }

  listGroup() {
    return this.groupsList
  }

  deleteGroup(groupId: string): void {
    this.groupsList = this.groupsList.filter(group => group.id !== groupId);
  }
}
