import * as React from 'react';
import { useEffect, useState } from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import type { IMyListFetchingProps } from './IMyListFetchingProps';
import MyListFetchingWebPart from '../MyListFetchingWebPart';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Icon, Label, Panel, Stack, TextField, PrimaryButton, DefaultButton } from '@fluentui/react';
import { Modal } from '@fluentui/react';
import { useBoolean } from '@fluentui/react-hooks';

const MyListFetching = (props: IMyListFetchingProps): JSX.Element => {

  interface User {
    Id: number;
    Title: string;
    FirstName: string;
    LastName: string;
    Country: string;
    Github: string;
  }

  const [users, setUsers] = useState<User[]>([]);
  const [editedUser, setEditedUser] = useState<User | null>(null);
  const [newUser, setNewUser] = useState<User>({ Id: 0, Title: '', FirstName: '', LastName: '', Country: '', Github: '' });
  const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);

  const handleEditUser = (item: User): void => {
    setEditedUser(item);
  };

  const handleDeleteUser = async (item: User): Promise<void> => {
    try {
      await MyListFetchingWebPart.sp.web.lists.getByTitle('philiplist').items.getById(item.Id).delete();
      const updatedUsers = users.filter(user => user.Id !== item.Id);
      setUsers(updatedUsers);
    } catch (error) {
      console.error("Error deleting user: ", error);
    }
  };

  const handleAddUser = async (): Promise<void> => {
    try {
      const addedUser = await MyListFetchingWebPart.sp.web.lists.getByTitle('philiplist').items.add({
        Country: newUser.Country,
        Github: newUser.Github
      });

      const updatedUsers = [...users, {
        Id: addedUser.data.Id,
        Title: addedUser.data.Title,
        FirstName: newUser.FirstName,
        LastName: newUser.LastName,
        Country: newUser.Country,
        Github: newUser.Github
      }];
      setUsers(updatedUsers);
      setNewUser({ Id: 0, Title: '', FirstName: '', LastName: '', Country: '', Github: '' });
      hideModal();
    } catch (error) {
      console.error("Error adding user: ", error);
    }
  };

  const handleSaveUser = async (): Promise<void> => {
    if (editedUser) {
      try {
        await MyListFetchingWebPart.sp.web.lists.getByTitle('philiplist').items.getById(editedUser.Id).update({
          Country: editedUser.Country,
          Github: editedUser.Github
        });

        const updatedUsers = users.map(user =>
          user.Id === editedUser.Id
            ? { ...user, Country: editedUser.Country, Github: editedUser.Github }
            : user
        );
        setUsers(updatedUsers);
        setEditedUser(null);
      } catch (error) {
        console.error("Error updating user: ", error);
      }
    }
  };

  useEffect(() => {
    const fetchUsers = async (): Promise<void> => {
      try {
        const items = await MyListFetchingWebPart.sp.web.lists.getByTitle('philiplist').items.expand('User').select('*,User/FirstName,User/LastName')();
        console.log('Fetched items:', items); // Debugging line

        const transformedItems = items.map(item => ({
          Id: item.Id,
          Title: item.Title,
          FirstName: item.User?.FirstName || '',
          LastName: item.User?.LastName || '',
          Country: item.Country || '',
          Github: item.Github || ''
        }));

        setUsers(transformedItems);
      } catch (error) {
        console.error("Error fetching users from SharePoint list: ", error);
      }
    };
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    fetchUsers();
  }, []);

  const columns: IColumn[] = [
    { key: 'column1', name: 'ID', fieldName: 'Id', minWidth: 100, maxWidth: 200, isResizable: true, onRender: (item) => <span>{item.Id}</span> },
    { key: 'column2', name: 'Name', fieldName: 'Name', minWidth: 100, maxWidth: 200, isResizable: true, onRender: (item) => <span>{`${item.FirstName} ${item.LastName}`}</span> },
    { key: 'column3', name: 'Country', fieldName: 'Country', minWidth: 100, maxWidth: 200, isResizable: true, onRender: (item) => <span>{item.Country}</span> },
    { key: 'column4', name: 'Github', fieldName: 'Github', minWidth: 100, maxWidth: 300, isResizable: true, onRender: (item) => <a href={item.Github} target="_blank" rel="noreferrer">Github</a> },
    {
      key: 'column5', name: ' ', fieldName: 'ID', minWidth: 100, maxWidth: 200, isResizable: true, onRender: (item) => <Stack horizontal tokens={{ childrenGap: 10 }}>
        <Icon iconName="Edit" onClick={() => handleEditUser(item)} />
        <Icon iconName="Delete" onClick={() => handleDeleteUser(item)} />
      </Stack>
    }
  ];

  return (
    <>
      <Stack onClick={showModal} horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
        <Label styles={{ root: { cursor: 'pointer' } }}>Add User</Label>
        <Icon iconName="Add" styles={{ root: { marginTop: '3px', cursor: 'pointer' } }} />
      </Stack>

      <Modal
        titleAriaId="Add User"
        isOpen={isModalOpen}
        onDismiss={hideModal}
        isBlocking={false}
        containerClassName="ms-modalExample-container"
        styles={{ main: { padding: '20px', width: '400px', borderRadius: '8px' } }}
      >
        <Stack tokens={{ childrenGap: 15 }} styles={{ root: { padding: '10px' } }}>
          <Label styles={{ root: { fontWeight: 'bold' } }}>Country</Label>
          <TextField
            value={newUser.Country}
            onChange={(e, newValue) => setNewUser({ ...newUser, Country: newValue || '' })}
            placeholder="Enter country"
          />

          <Label styles={{ root: { fontWeight: 'bold' } }}>Github</Label>
          <TextField
            value={newUser.Github}
            onChange={(e, newValue) => setNewUser({ ...newUser, Github: newValue || '' })}
            placeholder="Enter Github URL"
          />

          <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 20 } }}>
            <PrimaryButton text="Save" onClick={handleAddUser} />
            <DefaultButton text="Cancel" onClick={hideModal} />
          </Stack>
        </Stack>
      </Modal>

      <DetailsList
        columns={columns}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        setKey='set'
        items={users}
      />

      {!!editedUser && <Panel
        headerText="Edit User"
        isOpen={true}
        onDismiss={() => setEditedUser(null)}
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 15 }} styles={{ root: { padding: '10px' } }}>
          <Label>ID</Label>
          <TextField value={editedUser.Id.toString()} disabled />

          <Label>Name</Label>
          <TextField
            value={editedUser.FirstName + " " + editedUser.LastName}
            disabled
          />

          <Label>Country</Label>
          <TextField
            value={editedUser.Country}
            onChange={(e, newValue) => setEditedUser({ ...editedUser, Country: newValue || '' })}
          />

          <Label>Github</Label>
          <TextField
            value={editedUser.Github}
            onChange={(e, newValue) => setEditedUser({ ...editedUser, Github: newValue || '' })}
          />

          <PrimaryButton text="Save" onClick={handleSaveUser} styles={{ root: { marginTop: 20 } }} />
        </Stack>
      </Panel>}
    </>
  );
};

export default MyListFetching;
