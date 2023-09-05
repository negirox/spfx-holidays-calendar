import { SPFI, spfi, SPFx as spSPFx } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/views/list";
import { HolidaysListColumns, ListName } from "../constants/constant";

export class SPService {
	private sp: SPFI;
	constructor(context: any) {
		this.sp = spfi().using(spSPFx(context));
	}

	public getListItems = async (
		listTitle: string,
		filter: string = "",
		columns: string = "*",
		expand: string = "",
		orderby?: string,
		orderSequence?: boolean
	): Promise<any> => {
		let items: any = [];
		try {
			if (!orderby && filter !== "") {
				items = await this.sp.web.lists.getByTitle(listTitle).items.select(columns).filter(filter).expand(expand).top(5000)();
			}
			else if (orderby !== undefined && filter === "") {
				items = await this.sp.web.lists.getByTitle(listTitle).items.select(columns).orderBy(orderby, orderSequence).expand(expand).top(5000)();
			}
			else if (orderby === undefined && filter === "") {
				items = await this.sp.web.lists.getByTitle(listTitle).items.select(columns).expand(expand).top(5000)();
			}
			else {
				return await this.sp.web.lists.getByTitle(listTitle).items.select(columns).orderBy(orderby, orderSequence).filter(filter).expand(expand).top(5000)();
			}
			return Promise.resolve(items);
		} catch (ex) {
			return Promise.reject(ex);
		}
	};

	public getSharePointGroupDetails = async (groupTitle: string): Promise<any> => {
		try {
			const response = await this.sp.web.siteGroups.getByName(groupTitle)();
			return Promise.resolve(response);
		} catch (ex) {
			return Promise.reject(ex);
		}
	};

	public getUserDetails = async (userId: number): Promise<any> => {
		try {
			const response = await this.sp.web.getUserById(userId)();
			return Promise.resolve(response);
		} catch (ex) {
			return Promise.reject(ex);
		}
	};

	public ensureUser = async (loginName: string): Promise<any> => {
		if (loginName.indexOf("i:0#.f|membership|") === -1) {
			loginName = "i:0#.f|membership|" + loginName;
		}

		try {
			const response = await (await this.sp.web.ensureUser(loginName)).user();
			return Promise.resolve(response);
		} catch (ex) {
			return Promise.reject(ex);
		}
	};

	public async ensureHolidayList(): Promise<boolean> {
		const sp = this.sp;
		const _web = sp.web;
		let result = false;
	
		try {
		  await sp.web.lists.getByTitle(ListName.Holidays)().then(x => {
			console.log(ListName.Holidays);
		  }).catch(async x => {
	
			const ensureResult = await _web.lists.ensure(ListName.Holidays, "Holiday Calendar", 100, true);
			// if we've got the list
			if (ensureResult.list !== null) {
	
			  // if the list has just been created
			  if (ensureResult.created) {
				// we need to add the custom fields to the list
				await ensureResult.list.fields.addDateTime(HolidaysListColumns.Date);
				await ensureResult.list.fields.addText(HolidaysListColumns.Location, { MaxLength: 255, Required: false });
				await ensureResult.list.fields.addBoolean(HolidaysListColumns.Optional, { Required: false });
				const allItemsView = ensureResult.list.views.getByTitle('All Items');
				await allItemsView.fields.add(HolidaysListColumns.Date);
				await allItemsView.fields.add(HolidaysListColumns.Location);
				await allItemsView.fields.add(HolidaysListColumns.Optional);
				result = true;
			  }
			}
		  });
	
		} catch (e) {
		  // if we fail to create the list, write an exception in the _context log
		  result = false;
		}
	
		return result;
	  }
}
