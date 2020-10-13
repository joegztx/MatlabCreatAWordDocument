%word test
FilePath = fullfile(pwd,'Test\Template.docx'); 

actx_word = actxserver('Word.Application');
actx_word.Visible = true;
trace(actx_word.Visible);
word_handle = invoke(actx_word.Documents,'Open',FilePath);
actx_word.ActiveDocument.SaveAs2(fullfile(pwd,'Test\Report.docx'))

% Find end of document and make it the insertion point:
end_of_doc = get(actx_word.activedocument.content,'end');
set(actx_word.application.selection,'Start',end_of_doc);
set(actx_word.application.selection,'End',end_of_doc);

% Writing Train Data
actx_word.Selection.Style = 'Heading 1';
actx_word.Selection.TypeText('标题1');
actx_word.Selection.TypeParagraph;
actx_word.ActiveDocument.TablesOfContents.Item(1).Update;%update the content table

actx_word.Selection.Style = 'Heading 2';
actx_word.Selection.TypeText('标题2');
actx_word.Selection.TypeParagraph;

actx_word.Selection.TypeText('第一段落');
actx_word.Selection.TypeParagraph;
actx_word.Selection.TypeText('第二段落');
actx_word.Selection.TypeParagraph;
actx_word.Selection.TypeText('第三段落');
actx_word.Selection.TypeParagraph;

actx_word.Selection.Style = 'Heading 3'; 
actx_word.Selection.TypeText('标题3');
actx_word.Selection.TypeParagraph;

h=figure;
set(gcf,'position',[10 10 800 600]);
text(0.5,0.3,'text','FontSize',12,'HorizontalAlignment','center')
set(get(gcf,'CurrentAxes'),'Visible','off')
x=0:.5:10;
y=x.^2;
bar(x,y);
ylim([0,max(y)]);
xlabel('x');
ylabel('y');

text(max(xlim),min(ylim), {'ABB',''}, 'FontName', 'Arial', 'Color', 'red', 'FontSize', 11, 'HorizontalAlignment', 'right', 'VerticalAlignment', 'bottom')
text(max(xlim),min(ylim), {' ','Bordline'}, 'FontName', 'Arial', 'Color', [0.5,0.5,0.5], 'FontSize', 7, 'HorizontalAlignment', 'right', 'VerticalAlignment', 'bottom')
title('标题')
grid on

print(h,'-dmeta');
close(h);
invoke(actx_word.Selection,'Paste');
actx_word.Selection.TypeParagraph;

actx_word.Selection.Style = 'Heading 2';
actx_word.Selection.TypeText('标题4')
actx_word.Selection.TypeParagraph;

actx_word.Selection.TypeText('第四段落');
actx_word.Selection.TypeParagraph;

actx_word.Selection.InsertBreak;
actx_word.Selection.Style = 'Heading 1';
actx_word.Selection.TypeText('标题5')
actx_word.Selection.TypeParagraph;
actx_word.Selection.TypeText('第五段落');
actx_word.Selection.TypeParagraph;

%Insert a table
actx_word.Selection.Font.Size=10;
actx_word.ActiveDocument.Tables.Add(actx_word.Selection.Range,5,3);
actx_word.Selection.Font.Bold=1;
actx_word.Selection.TypeText('Table head 1');
actx_word.Selection.MoveRight;
actx_word.Selection.Font.Bold=1;
actx_word.Selection.TypeText('Table head 2');
actx_word.Selection.MoveRight;
actx_word.Selection.Font.Bold=1;
actx_word.Selection.TypeText('Table head 3');
actx_word.Selection.MoveRight;
actx_word.Selection.TypeParagraph;
actx_word.Selection.MoveDown;

actx_word.ActiveDocument.TablesOfContents.Item(1).Update
% actx_word.ActiveDocument.SaveAs2(fullfile(pwd,'\Report\Report.docx'))

delete(actx_word); 