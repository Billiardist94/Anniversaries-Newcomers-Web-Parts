export default interface RssItem {
    author: string;
    content: string;
    contentSnippet: string;
    id: string;
    link: string;
    pubDate?: Date;
    title: string;
    mediaURL?: string;
}