import { Component, inject } from '@angular/core';
import { GroupService } from '../group.service';
import { Group } from '../group';

@Component({
  selector: 'app-list-group',
  templateUrl: './list-group.component.html',
  styleUrl: './list-group.component.scss'
})
export class ListGroupComponent {
  groupsList: Group[] = []
  groupService: GroupService = inject(GroupService);

  getAllGroups() {
    this.groupsList = this.groupService.listGroup()
  }

  deleteGroup(groupId: string): void {
    this.groupService.deleteGroup(groupId)
    this.groupsList = this.groupService.listGroup()
  }
}
